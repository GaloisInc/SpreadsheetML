module Text.XML.SpreadsheetML.Builder where

import Text.XML.SpreadsheetML.Types

-- | Construct empty values
emptyWorkbook :: Workbook
emptyWorkbook = Workbook Nothing []

emptyDocumentProperties :: DocumentProperties
emptyDocumentProperties =
  DocumentProperties Nothing Nothing Nothing Nothing Nothing Nothing Nothing

emptyWorksheet :: Name -> Worksheet
emptyWorksheet name = Worksheet Nothing name

emptyTable :: Table
emptyTable =
  Table [] [] Nothing Nothing Nothing Nothing Nothing Nothing Nothing Nothing

emptyColumn :: Column
emptyColumn =
  Column Nothing Nothing Nothing Nothing Nothing Nothing

emptyRow :: Row
emptyRow = Row [] Nothing Nothing Nothing Nothing Nothing Nothing

emptyCell :: Cell
emptyCell = Cell Nothing Nothing Nothing Nothing Nothing

-- | Convenience constructors
number :: Double -> Cell
number d = emptyCell { cellData = Just (Number d) }

string :: String -> Cell
string s = emptyCell { cellData = Just (StringType s) }

bool :: Bool -> Cell
bool b = emptyCell { cellData = Just (Boolean b) }

-- | This function may change in future versions, if a real formula type is
-- created.
formula :: String -> Cell
formula f = emptyCell { cellFormula = Just (Formula f) }

mkWorkbook :: [Worksheet] -> Workbook
mkWorkbook ws = Workbook Nothing ws

mkWorksheet :: Name -> Table -> Worksheet
mkWorksheet name table = Worksheet (Just table) name

mkTable :: [Row] -> Table
mkTable rs = emptyTable { tableRows = rs }

mkRow :: [Cell] -> Row
mkRow cs = emptyRow { rowCells = cs }

-- | Most of the time this is the easiest way to make a table
tableFromCells :: [[Cell]] -> Table
tableFromCells cs = mkTable (map mkRow cs)
