import shutil
import datetime
from docx import Document
import glob
from docx2pdf import convert
from pdf2image import convert_from_path
import csv
import os

CSV_DIR = "csv/"
TEMPLATE_DOC = "document/template.docx"
DOC_DIR = "document/temp/"
PDF_DIR = "pdf/"
IMAGES_DIR = "images/"


def get_csv_data_list():
    """
    csvファイル一覧を取得し、値を配列で返す
    """

    csvFilePathList = glob.glob(CSV_DIR + "*.csv")

    csvDataList = []
    for csvFilePath in csvFilePathList:
        csvFile = open(csvFilePath, "r",
                       encoding="ms932", errors="", newline="")
        f = csv.reader(csvFile, delimiter=",", doublequote=True,
                       lineterminator="\r\n", quotechar='"', skipinitialspace=True)
        for row in f:
            csvDataList.append(row)

    return csvDataList


def create_document():
    """
    wordファイルを作成
    """

    now = datetime.datetime.now(datetime.timezone(
        datetime.timedelta(hours=9)))
    nowStr = now.strftime('%Y%m%d%H%M%S')

    csvDataList = get_csv_data_list()

    for row in csvDataList:
        if len(row) == 3:
            new_doc_path = DOC_DIR + row[2] + "_" + nowStr + ".docx"
            shutil.copyfile(TEMPLATE_DOC, new_doc_path)

            doc = Document(new_doc_path)

            # 行単位の置換
            for paragraph in doc.paragraphs:
                if paragraph.text == "${date}":
                    paragraph.text = row[0]
                if paragraph.text == "${reauesterName}":
                    paragraph.text = row[1]
                if paragraph.text == "${merchantName}":
                    paragraph.text = row[2]

            paragraphs = (paragraph
                          for table in doc.tables
                          for row in table.rows
                          for cell in row.cells
                          for paragraph in cell.paragraphs)

            # テーブルのセル単位の置換
            for paragraph in paragraphs:
                if paragraph.text == "${date}":
                    paragraph.text = row[0]
                if paragraph.text == "${reauesterName}":
                    paragraph.text = row[1]
                if paragraph.text == "${merchantName}":
                    paragraph.text = row[2]

            doc.save(new_doc_path)


def create_pdf():
    """
    wordファイルをpdfファイルに変換
    """

    docFilePathList = glob.glob(DOC_DIR + "*.docx")

    for docFilePath in docFilePathList:

        fileNameSplit = docFilePath.split("\\")
        fileName = fileNameSplit[len(fileNameSplit) - 1].split(".")[0]

        outputFile = PDF_DIR + fileName + ".pdf"
        file = open(outputFile, "w")
        file.close()

        convert(docFilePath, outputFile)


def create_image():
    """
    pdfファイルをjpegファイルに変換
    """
    pdfFilePathList = glob.glob(PDF_DIR + "*.pdf")

    for pdfFilePath in pdfFilePathList:
        fileNameSplit = pdfFilePath.split("\\")
        fileName = fileNameSplit[len(fileNameSplit) - 1].split(".")[0]

        pages = convert_from_path(pdfFilePath)
        for page in pages:
            page.save(IMAGES_DIR + fileName + '.jpg', 'JPEG')


shutil.rmtree(DOC_DIR)
os.mkdir(DOC_DIR)
shutil.rmtree(PDF_DIR)
os.mkdir(PDF_DIR)
shutil.rmtree(IMAGES_DIR)
os.mkdir(IMAGES_DIR)

create_document()
create_pdf()
create_image()
