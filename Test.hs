{-# LANGUAGE OverloadedStrings #-}
import Xlsx.Writer
import Xlsx.Sheet

main = saveXlsx "rcallahan" [("shump", [toRow (1::Int,2::Int)])] "test.xlsx"
