module Text.XML.SpreadsheetML.Writer where

import qualified Text.XML.SpreadsheetML.Types as T
import qualified Text.XML.Light as L
import qualified Text.XML.Light.Types as LT
import qualified Text.XML.Light.Output as O

import Control.Applicative ( (<$>) )
import Data.Maybe ( catMaybes, maybeToList )

--------------------------------------------------------------------------
-- | Convert a workbook to a string.  Write this string to a ".xls" file
-- and Excel will know how to open it.
showSpreadsheet :: T.Workbook -> String
showSpreadsheet wb = "<?xml version='1.0' ?>\n" ++
                     "<?mso-application progid=\"Excel.Sheet\"?>\n" ++
                     O.showElement (toElement wb)

---------------------------------------------------------------------------
-- | Namespaces
namespace   = L.blank_name { L.qURI    = Just "urn:schemas-microsoft-com:office:spreadsheet" }
oNamespace  = L.blank_name { L.qURI    = Just "urn:schemas-microsoft-com:office:office"
                           , L.qPrefix = Just "o" }
xNamespace  = L.blank_name { L.qURI    = Just "urn:schemas-microsoft-com:office:excel"
                           , L.qPrefix = Just "x" }
ssNamespace = L.blank_name { L.qURI    = Just "urn:schemas-microsoft-com:office:spreadsheet"
                           , L.qPrefix = Just "ss" }
htmlNamespace = L.blank_name { L.qURI = Just "http://www.w3.org/TR/REC-html40" }

--------------------------------------------------------------------------
-- | Empty Elements
emptyWorkbook :: LT.Element
emptyWorkbook = L.blank_element
  { L.elName    = workbookName
  , L.elAttribs = [xmlns, xmlns_o, xmlns_x, xmlns_ss, xmlns_html] }
  where
  workbookName = namespace { L.qName = "Workbook" }
  xmlns      = mkAttr "xmlns"      "urn:schemas-microsoft-com:office:spreadsheet"
  xmlns_o    = mkAttr "xmlns:o"    "urn:schemas-microsoft-com:office:office"
  xmlns_x    = mkAttr "xmlns:x"    "urn:schemas-microsoft-com:office:excel"
  xmlns_ss   = mkAttr "xmlns:ss"   "urn:schemas-microsoft-com:office:spreadsheet"
  xmlns_html = mkAttr "xmlns:html" "http://www.w3.org/TR/REC-html40"
  mkAttr k v = LT.Attr L.blank_name { L.qName = k } v

emptyDocumentProperties :: LT.Element
emptyDocumentProperties = L.blank_element { L.elName = documentPropertiesName }
  where
  documentPropertiesName = oNamespace { L.qName = "DocumentProperties" }

emptyWorksheet :: T.Name -> LT.Element
emptyWorksheet (T.Name n) = L.blank_element { L.elName    = worksheetName
                                            , L.elAttribs = [LT.Attr worksheetNameAttrName n] }
  where
  worksheetName = ssNamespace { L.qName   = "Worksheet" }
  worksheetNameAttrName = ssNamespace { L.qName   = "Name" }

emptyTable :: LT.Element
emptyTable = L.blank_element { L.elName = tableName }
  where
  tableName = ssNamespace { L.qName = "Table" }

emptyRow :: LT.Element
emptyRow = L.blank_element { L.elName = rowName }
  where
  rowName = ssNamespace { L.qName = "Row" }

emptyColumn :: LT.Element
emptyColumn = L.blank_element { L.elName = columnName }
  where
  columnName = ssNamespace { L.qName = "Column" }

emptyCell :: LT.Element
emptyCell = L.blank_element { L.elName = cellName }
  where
  cellName = ssNamespace { L.qName = "Cell" }

-- | Break from the 'emptyFoo' naming because you can't make
-- an empty data cell, except one holding ""
mkData :: T.ExcelValue -> LT.Element
mkData v = L.blank_element { L.elName     = dataName
                           , L.elContent  = [ LT.Text (mkCData v) ]
                           , L.elAttribs  = [ mkAttr v ] }
  where
  dataName   = ssNamespace { L.qName = "Data" }
  typeName s = ssNamespace { L.qName = s }
  typeAttr   = LT.Attr (typeName "Type")
  mkAttr (T.Number _)      = typeAttr "Number"
  mkAttr (T.Boolean _)     = typeAttr "Boolean"
  mkAttr (T.StringType _)  = typeAttr "String"
  mkCData (T.Number d)     = L.blank_cdata { LT.cdData = show d }
  mkCData (T.Boolean b)    = L.blank_cdata { LT.cdData = showBoolean b }
  mkCData (T.StringType s) = L.blank_cdata { LT.cdData = s }
  showBoolean True  = "1"
  showBoolean False = "0"

-------------------------------------------------------------------------
-- | XML Conversion Class
class ToElement a where
  toElement :: a -> LT.Element

-------------------------------------------------------------------------
-- | Instances
instance ToElement T.Workbook where
  toElement wb = emptyWorkbook
    { L.elContent = mbook ++
                    map (LT.Elem . toElement) (T.workbookWorksheets wb) }
    where
    mbook = maybeToList (LT.Elem . toElement <$> T.workbookDocumentProperties wb)

instance ToElement T.DocumentProperties where
  toElement dp = emptyDocumentProperties
    { L.elContent = map LT.Elem $ catMaybes
      [ toE T.documentPropertiesTitle       "Title"       id
      , toE T.documentPropertiesSubject     "Subject"     id
      , toE T.documentPropertiesKeywords    "Keywords"    id
      , toE T.documentPropertiesDescription "Description" id
      , toE T.documentPropertiesRevision    "Revision"    show
      , toE T.documentPropertiesAppName     "AppName"     id
      , toE T.documentPropertiesCreated     "Created"     id
      ]
    }
    where
    toE :: (T.DocumentProperties -> Maybe a) -> String -> (a -> String) -> Maybe L.Element
    toE fieldOf name toString = mkCData <$> fieldOf dp
      where
      mkCData cdata = L.blank_element
        { L.elName    = oNamespace { L.qName = name }
        , L.elContent = [LT.Text (L.blank_cdata { L.cdData = toString cdata })] }

instance ToElement T.Worksheet where
  toElement ws = (emptyWorksheet (T.worksheetName ws))
    { L.elContent = maybeToList (LT.Elem . toElement <$> (T.worksheetTable ws))  }

instance ToElement T.Table where
  toElement t = emptyTable
    { L.elContent = map LT.Elem $
      map toElement (T.tableColumns t) ++
      map toElement (T.tableRows t)
    , L.elAttribs = catMaybes
      [ toA T.tableDefaultColumnWidth  "DefaultColumnWidth"  show
      , toA T.tableDefaultRowHeight    "DefaultRowHeight"    show
      , toA T.tableExpandedColumnCount "ExpandedColumnCount" show
      , toA T.tableExpandedRowCount    "ExpandedRowCount"    show
      , toA T.tableLeftCell            "LeftCell"            show
      , toA T.tableFullColumns         "FullColumns"         showBoolean
      , toA T.tableFullRows            "FullRows"            showBoolean
      ] }
    where
    toA :: (T.Table -> Maybe a) -> String -> (a -> String) -> Maybe L.Attr
    toA fieldOf name toString = mkAttr <$> fieldOf t
      where
      mkAttr value = LT.Attr ssNamespace { L.qName = name } (toString value)

instance ToElement T.Row where
  toElement r = emptyRow
    { L.elContent = map LT.Elem $
      map toElement (T.rowCells r)
    , L.elAttribs = catMaybes
      [ toA T.rowCaption       "Caption"       showCaption
      , toA T.rowAutoFitHeight "AutoFitHeight" showAutoFitHeight
      , toA T.rowHeight        "Height"        show
      , toA T.rowHidden        "Hidden"        showHidden
      , toA T.rowIndex         "Index"         show
      , toA T.rowSpan          "Span"          show
      ] }

    where
    showAutoFitHeight T.AutoFitHeight      = "1"
    showAutoFitHeight T.DoNotAutoFitHeight = "0"
    toA :: (T.Row -> Maybe a) -> String -> (a -> String) -> Maybe L.Attr
    toA fieldOf name toString = mkAttr <$> fieldOf r
      where
      mkAttr value = LT.Attr ssNamespace { L.qName = name } (toString value)

showBoolean True  = "1"
showBoolean False = "0"

showCaption :: T.Caption -> String
showCaption (T.Caption s) = s

showHidden :: T.Hidden -> String
showHidden T.Hidden = "1"
showHidden T.Shown  = "0"

instance ToElement T.Column where
  toElement c = emptyColumn
    { L.elAttribs = catMaybes
      [ toA T.columnCaption      "Caption"      showCaption
      , toA T.columnAutoFitWidth "AutoFitWidth" showAutoFitWidth
      , toA T.columnHidden       "Hidden"       showHidden
      , toA T.columnIndex        "Index"        show
      , toA T.columnSpan         "Span"         show
      , toA T.columnWidth        "Width"        show
      ] }
    where
    showAutoFitWidth T.AutoFitWidth      = "1"
    showAutoFitWidth T.DoNotAutoFitWidth = "0"
    toA :: (T.Column -> Maybe a) -> String -> (a -> String) -> Maybe L.Attr
    toA fieldOf name toString = mkAttr <$> fieldOf c
      where
      mkAttr value = LT.Attr ssNamespace { L.qName = name } (toString value)

instance ToElement T.Cell where
  toElement c = emptyCell
    { L.elContent = map (LT.Elem . toElement) (maybeToList (T.cellData c))
    , L.elAttribs = catMaybes
      [ toA T.cellFormula     "Formula"     showFormula
      , toA T.cellIndex       "Index"       show
      , toA T.cellMergeAcross "MergeAcross" show
      , toA T.cellMergeDown   "MergeDown"   show
      ] }
    where
    showFormula (T.Formula f) = f
    toA :: (T.Cell -> Maybe a) -> String -> (a -> String) -> Maybe L.Attr
    toA fieldOf name toString = mkAttr <$> fieldOf c
      where
      mkAttr value = LT.Attr ssNamespace { L.qName = name } (toString value)

instance ToElement T.ExcelValue where
   toElement ev = mkData ev


