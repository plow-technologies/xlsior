{-# LANGUAGE OverloadedStrings #-}
module Xlsx.Writer where

import Codec.Archive.Zip
import Xlsx.Sheet
import Text.Blaze.Internal
import Text.Blaze.Renderer.Utf8
import Data.ByteString.Lazy (ByteString)
import Data.Text.Lazy (Text)
import qualified Data.Text as T

contTypes :: ByteString
contTypes = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"bin\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.customProperty\"/><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/><Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/><Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/><Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/></Types>"

customProp :: ByteString
customProp = "6\NUL3\NUL5\NUL4\NUL8\NUL8\NUL8\NUL6\NUL6\NUL3\NUL5\NUL0\NUL8\NUL7\NUL7\NUL0\NUL1\NUL7\NUL"

wbRels :: ByteString
wbRels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>"

appXml :: [T.Text] -> Markup
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
    step s = Append $ Parent "vt:lpstr" "<vt:lpstr" "</vt:lpstr>" (Content $ Text s)
