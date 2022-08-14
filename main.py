
import openpyxl as op
from openpyxl.drawing.image import Image
import os
import re


class AutoInjectionLogic():
    BASE_PATH = os.path.abspath(os.path.dirname(__file__))
    BASE_DIST_PATH = os.path.abspath(os.path.join(BASE_PATH, 'dist'))
    BASE_EXCEL = os.path.abspath(os.path.join(BASE_PATH, 'base.xlsx'))
    CELL_HEIGT = 18
    INCLEMENT_ROW = 3
    MAX_FILE_SIZE_MB = 50
    UNIT_BYTE = 1024

    def __init__(self) -> None:
        self.row = 1
        self.isEditing = False
        self.fileIndex = -1
        self.separateList = []
        self.onlyFile = True
        self.startName = ''
        self.midwayName = ''
        self.lastName = ''
        self.main()
        pass

    # サイズ変更
    def transformImageSize(self, image: Image):
        MAX_SIZE = 737
        if MAX_SIZE < image.height:
            ratio = MAX_SIZE / image.height
            image.height = MAX_SIZE
            image.width = int(image.width * ratio)

        if MAX_SIZE < image.width:
            ratio = MAX_SIZE / image.width
            image.width = MAX_SIZE
            image.height = int(image.height * ratio)
        return image

    # Excelのサイズを監視
    def watchFileCapacity(self, file):
        fileSize = os.path.getsize(file)
        kb = -(-fileSize // self.UNIT_BYTE)
        mb = -(-kb // self.UNIT_BYTE)
        return mb > self.MAX_FILE_SIZE_MB

    # 挿入する画像を配置しているディレクトリ選択
    def selectImageDir(self, path):
        self.row = 1
        files = os.listdir(path)
        for filename in files:
            if filename[-3:] == "png":
                fileUri = os.path.join(path, filename)
                img = Image(fileUri)
                img = self.transformImageSize(image=img)
                self.ws.add_image(img, 'B' + str(self.row))
                nextRow = -(-img.height // self.CELL_HEIGT)
                self.row = self.row + nextRow + self.INCLEMENT_ROW

    # Excelの修正開始
    def editExcel(self, path, sheetName):
        saveFile = os.path.join(self.distPath, self.excelName + '.xlsx')
        opneFile = saveFile if os.path.exists(saveFile) else self.BASE_EXCEL
        self.wb = op.load_workbook(opneFile)
        self.ws = self.wb.copy_worksheet(self.wb['Sheet1'])
        self.ws.title = sheetName
        self.selectImageDir(path)
        self.wb.save(saveFile)
        self.wb.close()
        if self.watchFileCapacity(saveFile):
            self.isEditing = False
            self.onlyFile = False
            self.midwayName = sheetName
            self.separateList.append((self.startName, self.midwayName))
            self.fileIndex = self.fileIndex + 1
            renameName = self.excelName + str(self.fileIndex)
            newSaveFile = os.path.join(self.distPath, renameName + '.xlsx')
            os.rename(saveFile, newSaveFile)

    def renameFiles(self):
        files = os.listdir(self.distPath)
        for index, file in enumerate(files):
            start, end = self.separateList[index]
            reName = self.excelName + '_' + start + '~' + end
            oldName = os.path.join(self.distPath, file)
            newName = os.path.join(self.distPath, reName + '.xlsx')
            os.rename(oldName, newName)

    def convertPathAndSheetName(self, path: str):
        path = os.path.dirname(path)
        sheetName = path.replace(self.imgDir, '').replace('\\', '-')
        sheetName = re.sub('^-', '', sheetName)
        return (path, sheetName)

    def findLastDir(self, dir=BASE_PATH, sheetList: list = []):
        dirList = os.listdir(dir)
        for dirOrFile in dirList:
            path = os.path.abspath(os.path.join(dir, dirOrFile))
            if os.path.isdir(path):
                self.findLastDir(dir=path)
            elif dirOrFile[-3:] == "png":
                path, sheetName = self.convertPathAndSheetName(path)
                if sheetName in sheetList:
                    continue
                if not self.isEditing:
                    self.startName = sheetName
                    self.isEditing = True
                sheetList.append(sheetName)
                self.editExcel(path, sheetName)
                self.lastName = sheetName

    # メイン関数
    def main(self):
        self.distPathName = input('Save Output Directory Name:')
        self.excelName = input('Save Excel Name : ')
        self.imgDir = input('Select Images Directory : ')
        self.distPath = os.path.join(self.BASE_DIST_PATH, self.distPathName)
        if not os.path.exists(self.distPath):
            os.makedirs(self.distPath)
        self.findLastDir(dir=self.imgDir)
        if not self.onlyFile:
            self.separateList.append((self.startName, self.lastName))
            self.renameFiles()


AutoInjectionLogic()
