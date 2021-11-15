import sys
import json
import subprocess
from pathlib import Path

def windows(paths, keep_active):
    import win32com.client

    word = win32com.client.Dispatch("Word.Application")
    wdFormatPDF = 17

    if paths["batch"]:
        for docx_filepath in sorted(Path(paths["input"]).glob("*.docx")):
            pdf_filepath = Path(paths["output"]) / (str(docx_filepath.stem) + ".pdf")
            doc = word.Documents.Open(str(docx_filepath))
            doc.SaveAs(str(pdf_filepath), FileFormat=wdFormatPDF)
            doc.Close()
    else:
        docx_filepath = Path(paths["input"]).resolve()
        pdf_filepath = Path(paths["output"]).resolve()
        doc = word.Documents.Open(str(docx_filepath))
        doc.SaveAs(str(pdf_filepath), FileFormat=wdFormatPDF)
        doc.Close()

    if not keep_active:
        word.Quit()


def macos(paths, keep_active):
    script = (Path(__file__).parent / "convert.jxa").resolve()
    cmd = [
        "/usr/bin/osascript",
        "-l",
        "JavaScript",
        str(script),
        str(paths["input"]),
        str(paths["output"]),
        str(keep_active).lower(),
    ]

    def run(cmd):
        process = subprocess.Popen(cmd, stderr=subprocess.PIPE)
        while True:
            line = process.stderr.readline().rstrip()
            if not line:
                break
            yield line.decode("utf-8")

    for line in run(cmd):
        try:
            msg = json.loads(line)
        except ValueError:
            continue
        if msg["result"] == "success":
            sys.exit(0)
        elif msg["result"] == "error":
            print(msg)
            sys.exit(1)


def resolve_paths(input_path, output_path):
    input_path = Path(input_path).resolve()
    output_path = Path(output_path).resolve() if output_path else None
    output = {}
    if input_path.is_dir():
        output["batch"] = True
        output["input"] = str(input_path)
        if output_path:
            assert output_path.is_dir()
        else:
            output_path = str(input_path)
        output["output"] = output_path
    else:
        output["batch"] = False
        assert str(input_path).endswith(".docx")
        output["input"] = str(input_path)
        if output_path and output_path.is_dir():
            output_path = str(output_path / (str(input_path.stem) + ".pdf"))
        elif output_path:
            assert str(output_path).endswith(".pdf")
        else:
            output_path = str(input_path.parent / (str(input_path.stem) + ".pdf"))
        output["output"] = output_path
    return output


def convert(input_path, output_path=None, keep_active=False):
    paths = resolve_paths(input_path, output_path)
    if sys.platform == "darwin":
        return macos(paths, keep_active)
    elif sys.platform == "win32":
        return windows(paths, keep_active)
    else:
        raise NotImplementedError(
            "docx2pdf requires Microsoft Word to be installed"
        )

