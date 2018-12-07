{-# LANGUAGE OverloadedStrings #-}
module Main where

import           Codec.Xlsx
import           Control.Lens
import           Jabara.Xlsx
import           System.Directory

main :: IO ()
main = do
  let sheet0 = foldr core0 def [RI 0 .. RI 99]
      sheet1 = foldr core1 def [CI 0 .. CI 19]
      book = def & atSheet "row orient" ?~ sheet0
                 & atSheet "column orient" ?~ sheet1
  -- let sheet = def & cellValueAt (cellTUnsafe "A1") ?~ CellText "foo"
  --     xlsx  = def & atSheet "List1" ?~ sheet
  home <- getHomeDirectory
  writeBookUnsafe (home ++ "/temp/temp.xlsx") book
  where
    core0 :: RowIndex -> Worksheet -> Worksheet
    core0 row sheet =
      let base = "A" +++ row
      in  sheet & cellValueAt base ?~ CellText "foo"
                & cellValueAt (base >>> "B") ?~ CellDouble 1.0
                & cellValueAt (base +>> 2) ?~ CellText "var"
    core1 :: ColumnIndex -> Worksheet -> Worksheet
    core1 col sheet =
      let base = tuple $ CellIndex (RI 0) col
      in  sheet & cellValueAt base ?~ CellText "foo"
                & cellValueAt (base ^^^ RI 1) ?~ CellDouble 1.0
                & cellValueAt (base +^^ 2) ?~ CellText "var"
