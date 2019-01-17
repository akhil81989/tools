class searchInExcel_lib():
	
    def __init__(self, _sheetList_var):
    	self._sheetList_var = _sheetList_var
        print self._sheetList_var
        self._readyToExportCellList = []
        
    def getCellIDContainingKeyWord(self,_firstCell, _lastCell, _checkKeyWords, _splitter = '*'):
        self._checkKeyWords = _checkKeyWords.split(_splitter)
        self.allCelList = self.listAllCellNames(_firstCell, _lastCell)
        foundList = []
        for eachCell in self.allCelList:
            tempSatisfiedCount = 0
            for eachSplitCheck in self._checkKeyWords:
                tempCellValue = self._sheetList_var[eachCell].value
                tempSearchResult = str(tempCellValue).find(eachSplitCheck)
                if tempSearchResult >= 0:
                    tempSatisfiedCount += 1
            if tempSatisfiedCount == len(self._checkKeyWords):
                foundList.append(eachCell)
        return foundList
        
    def listAllCellNames(self, _firstCell, _lastCell):
        _cellList = []
        tempFirstColoumn, tempFirstRow = self.seperateCellValue(_firstCell)
        tempLastColoumn, tempLastRow  = self.seperateCellValue(_lastCell)
        for tempColoumName in range(ord(tempFirstColoumn), ord(tempLastColoumn) + 1):
            for tempRowName in range(int(tempFirstRow), int(tempLastRow) + 1):
                _cellList.append(chr(tempColoumName) + str(tempRowName))
            print 'asdf'
        return _cellList
        
    def getAboveCellIfEmpty(self, _cellName):
        tempCellValueDivisionList = self.seperateCellValue(_cellName)
        tempCellRow= tempCellValueDivisionList[1]
        tempCellColoumn = tempCellValueDivisionList[0]
        while tempCellRow > 1:
            tempCurrCellName = tempCellColoumn + str(tempCellRow)
            tempReceivedValue = self._sheetList_var[tempCurrCellName].value
            if tempReceivedValue  != None:
                return tempCellColoumn + str(tempCellRow)
                break
            tempCellRow -=1
    
    def seperateCellValue(self, _cellValue):
        tempCombinedString = ''
        tempCombinedInt = ''
        for singleLetter in _cellValue:
            try:
                tempVar = int(singleLetter)
                tempCombinedInt += str(tempVar)
            except:
                tempCombinedString += singleLetter
        return [tempCombinedString, int(tempCombinedInt)]
        
        
    def getCombinedRow(self, _startingCell, _searchNext, _searchAboveForNone = True, _breakKeyWord = 'n/a', _filler = ''):
        tempCell = self.seperateCellValue(_startingCell)
        combinedRow = ''
        tempCorrectCellName = ''
        for tempCurrCellName in range(ord(tempCell[0]), ord(tempCell[0]) + _searchNext):
            tempCorrectCellName = self.getAboveCellIfEmpty(chr(tempCurrCellName) + str(tempCurrCellName))

            combinedRow += self._sheetList_var[tempCorrectCellName].value + _filler
        return combinedRow













