"""PPTX içindeki düzenlenebilir metni çevrimdışı İngilizceden Türkçeye çevirir."""

import argparse
import csv
from collections import Counter
from functools import lru_cache
import json
from pathlib import Path, PurePosixPath
import re
import shutil
import sys
from tempfile import TemporaryDirectory
from zipfile import BadZipFile, ZipFile

from lxml import etree


A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
PART = re.compile(
    r"ppt/(?:slides/slide|slideMasters/slideMaster|slideLayouts/slideLayout|"
    r"charts/chart)\d+\.xml"
)
NUMBER = re.compile(r"\d+(?:[.,:/-]\d+)*")
ACRONYM = re.compile(r"\b[A-Z]{2,}\d*\b")
ASCII_WORD = re.compile(r"[A-Za-z]{2,}")
PROTECTED = re.compile(r"https?://|www\.|\S+@\S+|\{[^{}]+\}|[A-Za-z]:\\", re.I)


def block_network(event, _args):
    if event.startswith("socket."):
        raise RuntimeError("Çeviri sırasında ağ erişimi engellendi.")


class LocalTranslator:
    def __init__(self, archive, directory):
        import ctranslate2
        import sentencepiece

        with ZipFile(archive) as package:
            for entry in package.namelist():
                name = PurePosixPath(entry)
                if name.name != "metadata.json":
                    continue
                metadata = json.loads(package.read(entry))
                if (metadata.get("from_code"), metadata.get("to_code")) == ("en", "tr"):
                    model_root = name.parent
                    break
            else:
                raise ValueError("Argos paketinde İngilizce → Türkçe modeli bulunamadı.")

            base = Path(directory)
            tokenizer = base / "sentencepiece.model"
            tokenizer.write_bytes(package.read(str(model_root / tokenizer.name)))
            model = base / "model"
            model.mkdir()
            for entry in package.namelist():
                name = PurePosixPath(entry)
                if (name.parent != model_root / "model" or entry.endswith("/")
                        or not re.fullmatch(r"[A-Za-z0-9_.-]+", name.name)
                        or name.name in {".", ".."}):
                    continue
                with package.open(entry) as source, (model / name.name).open("wb") as target:
                    shutil.copyfileobj(source, target)

        self.tokenizer = sentencepiece.SentencePieceProcessor(model_file=str(tokenizer))
        self.engine = ctranslate2.Translator(str(model), device="cpu")
        prefix = metadata.get("target_prefix")
        self.target_prefix = [[prefix]] if prefix else None

    def __call__(self, source):
        tokens = self.tokenizer.encode(source, out_type=str)
        if len(tokens) > 180:
            raise ValueError("Metin model için fazla uzun.")
        result = self.engine.translate_batch(
            [tokens], target_prefix=self.target_prefix, beam_size=4,
            length_penalty=0.2, max_input_length=0, max_decoding_length=260,
        )[0].hypotheses[0]
        if self.target_prefix and result[:1] == self.target_prefix[0]:
            result = result[1:]
        if len(result) >= 259 or "<unk>" in result:
            raise ValueError("Model sonucu eksik veya bilinmeyen sözcük içeriyor.")
        return self.tokenizer.decode_pieces(result).strip()


def check_translation(source, target):
    if not target or source.casefold() == target.casefold():
        return "Boş veya kaynakla aynı sonuç"
    if not any(ch.isalpha() for ch in target) or any(ord(ch) < 32 for ch in target):
        return "Geçersiz model çıktısı"
    if Counter(NUMBER.findall(source)) != Counter(NUMBER.findall(target)):
        return "Sayı veya sayısal gösterim değişti"
    # Tamamı büyük harf olan başlıklarda her kelimeyi kısaltma sayma.
    if not source.isupper():
        acronyms = Counter(ACRONYM.findall(source))
        if acronyms - Counter(ACRONYM.findall(target)):
            return "Kısaltma değişti veya kayboldu"
    if len(target) < len(source) * 0.33 or len(target) > max(18, len(source) * 2):
        return "Çeviri uzunluğu olağan aralığın dışında"
    return None


def translate_paragraph(paragraph, translate):
    runs = paragraph.findall(A + "r")
    nodes = [run.find(A + "t") for run in runs]
    nodes = [node for node in nodes if node is not None]
    source = "".join(node.text or "" for node in nodes).strip()
    if not source:
        return False, None
    if not ASCII_WORD.search(source) or re.fullmatch(r"[A-Z0-9_-]{2,3}", source):
        return False, "Sayı, kısa kısaltma veya çevrilebilir sözcük yok"
    if PROTECTED.search(source):
        return False, "Bağlantı, e-posta, dosya yolu veya yer tutucu içeriyor"
    if paragraph.find(A + "fld") is not None or paragraph.find(A + "br") is not None:
        return False, "Dinamik alan veya elle eklenmiş satır sonu içeriyor"
    if any(ch in source for ch in "\n\r\t"):
        return False, "Gömülü satır sonu veya sekme içeriyor"

    styles = set()
    for node in nodes:
        if node.text:
            props = node.getparent().find(A + "rPr")
            styles.add(etree.tostring(props, with_tail=False) if props is not None else b"")
    if len(styles) > 1:
        return False, "Paragraf içinde farklı karakter biçimleri var"

    try:
        target = translate(source)
    except ValueError:
        return False, "Model sonucu kullanılamadı veya metin fazla uzun"
    except Exception as error:
        return False, f"Çeviri hatası ({type(error).__name__})"
    reason = check_translation(source, target)
    if reason:
        return False, reason

    original = "".join(node.text or "" for node in nodes)
    first = next(node for node in nodes if node.text)
    for node in nodes:
        node.text = ""
    leading = original[:len(original) - len(original.lstrip())]
    trailing = original[len(original.rstrip()):]
    first.text = leading + target + trailing
    return True, None


def translate_deck(source, destination, translate):
    rows = []
    updated = 0
    parser = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)

    with ZipFile(source) as original, ZipFile(destination, "w") as output:
        names = original.namelist()
        if "ppt/presentation.xml" not in names or len(set(names)) != len(names):
            raise ValueError("Geçerli bir PowerPoint paketi bulunamadı.")
        output.comment = original.comment
        for entry in original.infolist():
            name = entry.filename
            data = original.read(entry)
            if PART.fullmatch(name):
                tree = etree.fromstring(data, parser).getroottree()
                if tree.docinfo.doctype:
                    raise ValueError("DTD içeren XML parçası desteklenmiyor.")
                changed = False
                for index, paragraph in enumerate(tree.iter(A + "p"), 1):
                    # Ana slayt ve düzenlerdeki varsayılan PowerPoint
                    # yer tutucu talimatlarını düzenleme.
                    if name.startswith(("ppt/slideMasters/", "ppt/slideLayouts/")):
                        shape = next(paragraph.iterancestors(P + "sp"), None)
                        if shape is not None and shape.find(".//" + P + "ph") is not None:
                            continue
                    success, reason = translate_paragraph(paragraph, translate)
                    if reason:
                        rows.append((name, index, reason))
                    if success:
                        updated += 1
                        changed = True

                if name.startswith("ppt/charts/") and any(
                    list(tree.iter(C + tag)) for tag in ("strRef", "strLit", "multiLvlStrRef")
                ):
                    rows.append((name, "", "Veri kaynağına bağlı grafik etiketleri atlandı"))
                if name.startswith("ppt/slides/") and next(tree.iter(P + "pic"), None) is not None:
                    rows.append((name, "", "Görsel var. Görsel içi yazılar taranmadı"))
                if changed:
                    data = etree.tostring(tree, encoding="UTF-8", xml_declaration=True)
            elif name.startswith("ppt/diagrams/data") and name.endswith(".xml"):
                rows.append((name, "", "SmartArt metni atlandı"))
            output.writestr(entry, data)

    return updated, rows


def main():
    cli = argparse.ArgumentParser(description=__doc__)
    cli.add_argument("sunum", type=Path, help="Kaynak PPTX dosyası")
    cli.add_argument("--model", type=Path, required=True, help="Yerel İngilizce → Türkçe .argosmodel")
    cli.add_argument("--cikti", type=Path, help="Çıktı PPTX yolu. Varsayılan: kaynak_adı_tr.pptx")
    args = cli.parse_args()
    source = args.sunum.resolve()
    target = (args.cikti or source.with_name(source.stem + "_tr.pptx")).resolve()
    report = target.with_suffix(".atlananlar.csv")
    if source.suffix.lower() != ".pptx" or target.suffix.lower() != ".pptx":
        cli.error("Kaynak ve çıktı uzantıları .pptx olmalı.")
    if target == source or target.exists() or report.exists():
        cli.error("Çıktı veya rapor yolu kullanımda. --cikti ile farklı bir ad belirtin.")

    sys.addaudithook(block_network)
    try:
        with TemporaryDirectory(prefix="pptx_model_") as model_dir:
            engine = LocalTranslator(args.model, model_dir)
            if check_translation("This is a test.", engine("This is a test.")):
                raise RuntimeError("Yerel model başlangıç kontrolünü geçemedi.")

            with TemporaryDirectory(prefix=".pptx_tr_", dir=target.parent) as staging:
                staged_pptx = Path(staging) / "output.pptx"
                count, rows = translate_deck(source, staged_pptx, lru_cache(2048)(engine))
                with ZipFile(staged_pptx) as archive:
                    if archive.testzip() is not None:
                        raise ValueError("Oluşturulan PPTX arşivi doğrulanamadı.")
                staged_report = Path(staging) / "report.csv"
                with staged_report.open("w", encoding="utf-8-sig", newline="") as stream:
                    writer = csv.writer(stream, delimiter=";")
                    writer.writerow(("PPTX parçası", "Paragraf sırası", "Açıklama"))
                    writer.writerows(rows)
                # Windows'ta rename var olan hedefin üzerine yazmaz.
                staged_report.rename(report)
                try:
                    staged_pptx.rename(target)
                except OSError:
                    report.unlink()
                    raise
    except (BadZipFile, KeyError, OSError, RuntimeError, ValueError, etree.XMLSyntaxError) as error:
        cli.exit(1, f"İşlem tamamlanamadı ({type(error).__name__}). {error}\n")

    print(f"Güncellenen paragraf: {count}")
    print(f"Sunum: {target}")
    print(f"Kontrol raporu: {report}")


if __name__ == "__main__":
    main()
