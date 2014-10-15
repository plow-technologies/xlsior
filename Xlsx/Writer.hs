{-# LANGUAGE OverloadedStrings #-}
module Xlsx.Writer where

import Codec.Archive.Zip
import Data.Monoid
import Xlsx.Sheet
import Text.Blaze.Internal
import Text.Blaze.Renderer.Utf8
import Data.ByteString.Lazy (ByteString)
import qualified Data.ByteString.Lazy as L
import Data.Text.Lazy (Text)
import Data.Time.Format
import Data.Time.Clock.POSIX
import Data.Time.Clock
import System.Locale
import qualified Data.Text as T

contentTypes :: Int -> Markup
contentTypes n = Append decl $ types $ Append (Content $ Static "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/><Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/><Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>") sheets where
    sheets = mconcat $ map sheet [1..n]
    types = AddAttribute "xmlns" " xmlns=\"" "http://schemas.openxmlformats.org/package/2006/content-types" .
        Parent "Types" "<Types" "</Types>"
    conttype = AddAttribute "ContentType" " ContentType=\""
    sheet :: Int -> Markup
    sheet n = conttype "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" $ 
        AddAttribute "PartName" " PartName=\"" (String $ "/xl/worksheet/sheet" ++ show n ++ ".xml") $
        Parent "Override" "<Override" "</Override>" Empty

rootRelXml :: ByteString
rootRelXml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/></Relationships>"

appXml :: ByteString
appXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><TotalTime>0</TotalTime></Properties>"

{-appXml :: [T.Text] -> Markup
appXml sheets = Append decl $
    AddAttribute "xmlns" " xmlns=\"" "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" $
    AddAttribute "xmlns:vt" " xmlns:vt=\"" "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" $
    Parent "Properties" "<Properties" "</Properties>" $
    Append (Content $ Static "<Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size=\"2\" baseType=\"variant\"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>") $
    Append (Parent "TitlesOfParts" "<TitlesOfParts" "</TitlesOfParts>" $
    AddAttribute "baseType" " baseType=\"" "lpstr" $
    AddAttribute "size" " size=\"" (String $ show $ length sheets) $
    Parent "vt:vector" "<vt:vector" "</vt:vector>" sheetMarkup) $
    Content $ Static "<LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>14.0300</AppVersion></Properties>" where
    sheetMarkup = foldr step Empty sheets
    step s = Append $ Parent "vt:lpstr" "<vt:lpstr" "</vt:lpstr>" (Content $ Text s) -}

coreXml :: FormatTime t => T.Text -> t -> Markup
coreXml name time = Append decl $
    AddAttribute "xmlns:cp" " xmlns:cp=\"" "http://schemas.openxmlformats.org/package/2006/metadata/core-properties" $
    AddAttribute "xmlns:dc" " xmlns:dc=\"" "http://purl.org/dc/elements/1.1/" $
    AddAttribute "xmlns:dcterms" " xmlns:dcterms=\"" "http://purl.org/dc/terms/" $
    AddAttribute "xmlns:dcmitype" " xmlns:dcmitype=\"" "http://purl.org/dc/dcmitype/" $
    AddAttribute "xmlns:xsi" " xmlns:xsi=\"" "http://www.w3.org/2001/XMLSchema-instance" $
    Parent "cp:coreProperties" "<cp:coreProperties" "</cp:coreProperties>" $
    Append (Parent "dc:creator" "<dc:creator" "</dc:creator>" name') $
    Append (Parent "cp:lastModifiedBy" "<cp:lastModifiedBy" "</cp:lastModifiedBy>" name') $
    Append (w3cdtf $ Parent "dcterms:created" "<dcterms:created" "</dcterms:created>" time') $
    w3cdtf $ Parent "dcterms:modified" "<dcterms:modified" "</dcterms:modified>" time' where
    w3cdtf = AddAttribute "xsi:type" " xsi:type=\"" "dcterms:W3CDTF"
    name' = Content $ Text name
    time' = Content $ Text $ T.pack $ formatTime defaultTimeLocale "%Y-%m-%dT%H:%M:%SZ" time

theme1 :: Markup
theme1 = decl

workbook :: [T.Text] -> Markup
workbook sheets = Append decl $
    AddAttribute "xmlns" " xmlns=\"" "http://schemas.openxmlformats.org/spreadsheetml/2006/main" $
    AddAttribute "xmlns:r" " xmlns:r=\"" "http://schemas.openxmlformats.org/officeDocument/2006/relationships" $
    Parent "workbook" "<workbook" "</workbook>" $ Parent "sheets" "<sheets" "</sheets>" $ sheets' where
    sheets' = mconcat $ map sheettag $ zip [1..] sheets
    sheettag :: (Int, T.Text) -> Markup
    sheettag (n, s) = AddAttribute "name" " name=\"" (Text s) $
        AddAttribute "sheetId" " sheetId=\"" (String $ show n) $
        AddAttribute "r:id" " r:id=\"" (String $ "rId" ++ show n) $
        Parent "sheet" "<sheet" "</sheet>" Empty

workbookRels :: Int -> Markup
workbookRels sheets = Append decl $
    AddAttribute "xmlns" " xmlns=\"" "http://schemas.openxmlformats.org/package/2006/relationships" $
    Parent "Relationships" "<Relationships" "</Relationships>" $ Append sheets' otherrels where
        sheets' = mconcat $ map sheetrel [1..sheets]
        typeattr = AddAttribute "Type" " Type=\""
        styles = typeattr "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
        rId n = AddAttribute "Id" " Id=\"" (String $ "rId" ++ show n)
        target = AddAttribute "Target" " Target=\""
        sheetrel :: Int -> Markup
        sheetrel n = typeattr "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" $
            rId n $ target (String $ "worksheets/sheet" ++ show n ++ ".xml") reltag
        reltag = Parent "Relationship" "<Relationship" "</Relationship>" Empty
        otherrels = rId (sheets+1) $ target "styles.xml" $
            typeattr "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" reltag

styles :: ByteString
styles = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"></styleSheet>"

sheetsLBS :: UTCTime -> T.Text -> [(T.Text, [Row])] -> ByteString
sheetsLBS time creator sheets = fromArchive $ foldr addEntryToArchive emptyArchive entries where
    t = round $ utcTimeToPOSIXSeconds time
    entries = [toEntry "[Content_Types].xml" t (renderMarkup $ contentTypes $ length sheets), 
        toEntry "_rels/.rels" t rootRelXml,
        toEntry "docProps/app.xml" t appXml,
        toEntry "docProps/core.xml" t (renderMarkup $ coreXml creator time),
        toEntry "xl/styles.xml" t styles,
        toEntry "xl/workbook.xml" t (renderMarkup $ workbook $ map fst sheets),
        toEntry "xl/theme/theme1.xml" t (renderMarkup theme1),
        toEntry "xl/_rels/workbook.xml.rels" t (renderMarkup $ workbookRels $ length sheets)] ++ sheets'
    sheets' = map (\(n,(_,rs)) -> toEntry ("xl/worksheets/sheet" ++ show n ++ ".xml") t (renderMarkup $ renderSheet rs)) $
        zip [1..] sheets

saveXlsx :: T.Text -> [(T.Text, [Row])] -> FilePath -> IO ()
saveXlsx creator sheets path = do
    curtime <- getCurrentTime
    L.writeFile path $ sheetsLBS curtime creator sheets
