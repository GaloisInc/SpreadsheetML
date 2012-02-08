module Text.XML.SpreadsheetML.Types where

{- See http://msdn.microsoft.com/en-us/library/aa140066%28office.10%29.aspx -}

import Data.Word ( Word64 )

-- | Only implement what we need

data Workbook = Workbook
  { workbookDocumentProperties :: Maybe DocumentProperties
--  , workbookExcelWorkbook      :: Maybe ExcelWorkbook
  , workbookWorksheets         :: [Worksheet]
--  , workbookStyles             :: [Style]
  }
  deriving (Read, Show)

data DocumentProperties = DocumentProperties
  { documentPropertiesTitle       :: Maybe String
  , documentPropertiesSubject     :: Maybe String
  , documentPropertiesKeywords    :: Maybe String
  , documentPropertiesDescription :: Maybe String
  , documentPropertiesRevision    :: Maybe Word64
  , documentPropertiesAppName     :: Maybe String
  , documentPropertiesCreated     :: Maybe String -- ^ Actually, this should be a date time
  }
  deriving (Read, Show)

data Worksheet = Worksheet
  { worksheetTable       :: Maybe Table
  , worksheetName        :: Name
  }
  deriving (Read, Show)

data Table = Table
  { tableColumns             :: [Column]
  , tableRows                :: [Row]
  , tableDefaultColumnWidth  :: Maybe Double -- ^ Default is 48
  , tableDefaultRowHeight    :: Maybe Double -- ^ Default is 12.75
  , tableExpandedColumnCount :: Maybe Word64
  , tableExpandedRowCount    :: Maybe Word64
  , tableLeftCell            :: Maybe Word64 -- ^ Default is 1
--  , tableStyleID             :: Maybe StyleID
  , tableTopCell             :: Maybe Word64 -- ^ Default is 1
  , tableFullColumns         :: Maybe Bool
  , tableFullRows            :: Maybe Bool
  }
  deriving (Read, Show)

data Column = Column
  { columnCaption      :: Maybe Caption
  , columnAutoFitWidth :: Maybe AutoFitWidth
  , columnHidden       :: Maybe Hidden
  , columnIndex        :: Maybe Word64
  , columnSpan         :: Maybe Word64
--  , columnStyleId      :: Maybe StyleID
  , columnWidth        :: Maybe Double
  }
  deriving (Read, Show)

data Row = Row
  { rowCells         :: [Cell]
  , rowCaption       :: Maybe Caption
  , rowAutoFitHeight :: Maybe AutoFitHeight
  , rowHeight        :: Maybe Double
  , rowHidden        :: Maybe Hidden
  , rowIndex         :: Maybe Word64
  , rowSpan          :: Maybe Word64
--  , rowStyleID       :: Maybe StyleID
  }
  deriving (Read, Show)

data Cell = Cell
  -- elements
  { cellData          :: Maybe ExcelValue
  -- Attributes
  , cellFormula       :: Maybe Formula
  , cellIndex         :: Maybe Word64
  , cellMergeAcross   :: Maybe Word64
  , cellMergeDown     :: Maybe Word64
--  , cellStyleID       :: Maybe StyleID
  }
  deriving (Read, Show)

data ExcelValue = Number Double | Boolean Bool | StringType String
  deriving (Read, Show)

-- | TODO: Currently just a string, but we could model excel formulas and
-- use that type here instead.
newtype Formula = Formula String
  deriving (Read, Show)

data AutoFitWidth = AutoFitWidth | DoNotAutoFitWidth
  deriving (Read, Show)

data AutoFitHeight = AutoFitHeight | DoNotAutoFitHeight
  deriving (Read, Show)

-- | Attribute for hidden things
data Hidden = Shown | Hidden
  deriving (Read, Show)

-- | For now this is just a string, but we could model excel's names
newtype Name = Name String
  deriving (Read, Show)

newtype Caption = Caption String
  deriving (Read, Show)

-----------------------------------
-- Hopefully all this down here is uncessary
{-
-- TODO: do we need this type?
{-
data ExcelWorkbook = ExcelWorkbook
  { excelWorkbookWindowHeight :: Maybe Word64
  , excelWorkbookWindowWidth  :: Maybe Word64
  , excelWorkbookWindowTopX   :: Maybe Word64
  , excelWorkbookWindowTopY   :: Maybe Word64
  }
-}

{-
data Style = Style
  { styleAlignment    :: Maybe Alignment
  , styleBorders      :: Maybe Borders
  , styleFont         :: Maybe Font
  , styleInterior     :: Maybe Interior
  , styleNumberFormat :: Maybe NumberFormat
  , styleProtection   :: Maybe Protection
  , styleID           :: StyleID
  , styleName         :: Maybe Name
  , styleParent       :: Maybe StyleID
  }
  deriving (Read, Show)
-}

-- | For now, the value is the name of the style to use
{-
newtype StyleID = StyleID String
  deriving (Read, Show)
-}

-- | Bold tag
data Bold = Bold
  deriving (Read, Show)

-- | This is not used by excel, but it is still part of the XML schema
data ComponentOptions = ComponentOptions
  { componentOptionsToolbar  :: (Maybe Toolbar)
  }
  deriving (Read, Show)

-- | Having this tag means that non-standard row/column headers should be used
data DisplayCustomHeaders = DisplayCustomHeaders
  deriving (Read, Show)

-- | If present, excel should hide the office logo
data HideOfficeLogo = HideOfficeLogo
  deriving (Read, Show)

-- | Configure the office toolbar status
-- Default vaule for HiddenAttr is shown
data Toolbar = Toolbar
  { toolbarHideOfficeLogo :: (Maybe HideOfficeLogo)
  , toolbarHidden         :: (Maybe Hidden)
  }
  deriving (Read, Show)

-- | Options specific to the spreadsheet
data WorksheetOptions = WorksheetOptions
  { worksheetOptionsDisplayCustomHeaders :: (Maybe DisplayCustomHeaders)
  }
  deriving (Read, Show)

-- | Font tag
newtype Font = Font
  { fontColor :: (Maybe Color)
  }
  deriving (Read, Show)

newtype Color = Color String -- ^ Name of the color to use
  deriving (Read, Show)

-- Italics tag
data Italic = Italic
  deriving (Read, Show)

type SmartTags = [SmartTag]

-- | SmartTag contains the name and the namespace uri
data SmartTagType = SmartTagType
  { smartTagTypeName         :: String
  , smartTagTypeNamespaceURI :: String
  }
  deriving (Read, Show)

data SmartTag = SmartTag
  deriving (Read, Show)

-- | Strikethrough tag
data Strikethrough = Strikethrough
  deriving (Read, Show)

-- | For outline
-- TODO: The documentation about this element seems to be wrong, it says
-- it has no optional attributes and then gives an example using the ss:Style
-- attribute
data Span = Span
  deriving (Read, Show)

-- | Alignment attributes for a style
data Alignment = Alignment
  { alignmentHorizontal   :: Maybe Horizontal
--  , alignmentIndent       :: Maybe Word64        -- Not supported in SpreadsheetML
  , alignmentReadingOrder :: Maybe ReadingOrder
--  , alignmentRotate       :: Maybe Double        -- Not supported in SpreadsheetML
--  , alignmentShrinkToFit  :: Maybe ShrinkToFit   -- Not supported in SpreadsheetML
  , alignmentVertical     :: Maybe Vertical
--  , alignmentVerticalText :: Maybe VerticalText  -- Not supported in SpreadsheetML
--  , alignmentWrapText     :: Maybe WrapText      -- Not supported in SpreadsheetML
  }
  deriving (Read, Show)

-- | Spreadsheets only support Automatic (default), Left, Centor, and Right.
-- The rest are here for completeness.
data Horizontal = HAutomatic | HLeft | HCenter | HRight
  deriving (Read, Show)
{-
                | Fill | Justify | CenterAcrossSelection
                | Distributed | JustifyDistributed
-}

-- | Default is context, which isn't supported.  That seems like a documentation bug
data ReadingOrder = RightToLeft | LeftToRight
  deriving (Read, Show)
--                  | Context -- not supported in SpreadsheetML

-- | not supported in SpreadsheetML
data ShrinkToFit = ShrinkToFit | DoNotShrinkToFit
  deriving (Read, Show)

data Vertical = VAutomatic | VTop | VBottom | VCenter
  deriving (Read, Show)
--              | Justify | Distributed | JustifyDistributed -- Not supported

-- | Not supported in SpreadsheetML
data VerticalText = VerticalText | HorizontalText
  deriving (Read, Show)

-- | Not suppored in SpreadsheetML
data WrapText = WrapText | DoNotWrapText
  deriving (Read, Show)

data Border = Border
  { borderPosition  :: Position
  , borderColor     :: Maybe Color
  , borderLineStyle :: Maybe LineStyle
  , borderWeight    :: Maybe Weight
  }
  deriving (Read, Show)

data Position = PLeft | PTop | PRight | PBottom
  deriving (Read, Show)
--              | PDiagonalLeft | PDiagonalRight -- Not supported in SpreadsheetML

data LineStyle = LineStyleNone | Continuous | Dash | Dot | DashDot | DashDotDot
  deriving (Read, Show)
--               | SlantDashDot | LSDouble  -- Not supported in SpreadsheetML

-- | Hairline = 0, Thin = 1, Medium = 2, Thick = 3
data Weight = Hairline | Thin | Medium | Thick
  deriving (Read, Show)

newtype Borders = Borders [Border]
  deriving (Read, Show)

-- | TODO: Currently just a String but we could model excel references and
-- use that type instead.
newtype ArrayRange = ArrayRange String
  deriving (Read, Show)

data Comment = Comment
  { commentAuthor     :: Maybe String
  , commentShowAlways :: Maybe ShowAlways
  }
  deriving (Read, Show)

data ShowAlways = ShowAlways | DoNotShowAlways
  deriving (Read, Show)

data Ticked = Ticked | NotTicked
  deriving (Read, Show)

data FontStyle = FontStyle
  { fontStyleBold          :: Maybe Bold    -- ^ Default is non-bold
  , fontStyleColor         :: Maybe String  -- ^ Default is "automatic"
  , fontStyleName          :: Maybe String  -- ^ Default is "Arial"
  , fontStyleItalic        :: Maybe Italic  -- ^ Default is non-italic
--  , fontStyleOutline       :: Maybe Outline -- ^ Default is non-outline  -- Not supported in SpreadsheetML
--  , fontStyleShadow        :: Maybe Bool -- ^ Default is false, not supported in SpreadsheetML
  , fontStyleSize          :: Maybe Double  -- ^ Default is 10
  , fontStyleStrikethrough :: Maybe Strikethrough -- ^ Default is non-strikethrough
  , fontStyleUnderline     :: Maybe Underline -- ^ Default is non-underline
--  , fontVerticalAlign      :: Maybe VerticalAlign -- ^ Default is None, not supported in SpreadsheetML
  , fontStyleCharSet       :: Maybe Word64 -- ^ Windows specific, defaults to 0
  , fontStyleFamily        :: Maybe Family -- ^ Default is Automatic
  }
  deriving (Read, Show)

-- TODO: Add something for Strikethrough type
-- TODO: Add something for underline type

data Family = Automatic | Decorative | Modern | Roman | Script | Swiss
  deriving (Read, Show)

data Interior = Interior
  { interiorColor        :: Maybe Color   -- ^ Default is "Automatic"
  , interiorPattern      :: Maybe Pattern -- ^ Default is "None"
  , interiorPatternColor :: Maybe Color   -- ^ Default is "Automatic"
  }
  deriving (Read, Show)

data Pattern = PatternNone | Solid | Gray75 | Gray50 | Gray25 | Gray125
             | Gray0625 | HorizStripe | VertStripe | ReverseDiagStripe
             | DiagStripe | DiagCross | ThickDiagCross | ThinHorzStripe
             | ThinVertStripe | ThinReverseDiagStripe | ThinDiagStripe
             | ThinHorzCross | ThinDiagCross
  deriving (Read, Show)

data NamedCell = NamedCell
  { namedCellName :: Name
  }
  deriving (Read, Show)

data NamedRange = NamedRange
  { namedRangeName     :: Name
  , namedRangeRefersTo :: RefersTo
  , namedRangeHidden   :: Maybe Hidden
  }
  deriving (Read, Show)

-- | For now this is just a string, but we could model excel's formulas
newtype RefersTo = RefersTo String
  deriving (Read, Show)

type Names = [NamedRange]

data NumberFormat = NumberFormat
  { numberFormat :: Format
  }
  deriving (Read, Show)

data Format = Custom String
            | General | GeneralNumber | GeneralDate | LongDate
            | MediumDate | ShortDate | LongTime | MediumTime | ShortTime
            | Currency | EuroCurrency | Fixed | Standard | Percent
            | Scientific | YesNo | TrueFalse | OnOff
  deriving (Read, Show)

data Protection = Protection
  { protectionProtected   :: Maybe Protected
  , protectionHideFormula :: Maybe HideFormula
  }
  deriving (Read, Show)

data Protected   = Protected   | NotProtected
  deriving (Read, Show)
data HideFormula = HideFormula | DoNotHideFormula
  deriving (Read, Show)

data Sub = Sub
  deriving (Read, Show)

data Sup = Sup
  deriving (Read, Show)

data Underline = Underline (Maybe UnderlineStyle)
  deriving (Read, Show)

data UnderlineStyle = Single | DoubleLine | SingleAccounting | DoubleAccounting
  deriving (Read, Show)

data AutoFilter = AutoFilter
  deriving (Read, Show)
{-  TODO: Finish this type and related types
data AutoFilter = AutoFilter
  { autoFilterColumn :: Maybe AutoFilterColumn
  , autoFilterRange  :: Range
  }
  deriving (Read, Show)
-}
-- | TODO: Implement a better type for Excel ranges
newtype Range = Range String
  deriving (Read, Show)

-- TODO: Implement this
data AutoFilterAnd = AutoFilterAnd
  deriving (Read, Show)

data WorksheetOptionsX = WorksheetOptionsX
  deriving (Read, Show)
{- TODO: Finish this and related types
data WorksheetOptionsX = WorksheetOptionsX
  { worksheetOptionsPageSetup :: Maybe PageSetup
  }
  deriving (Read, Show)
-}


-}
