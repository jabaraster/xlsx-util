{-# LANGUAGE OverloadedStrings #-}
module Jabara.Xlsx.TypeSpec where

import           Jabara.Xlsx.Types
import           Test.Hspec

spec :: Spec
spec = do
    it "parse column index text" $ do
      parseColumnIndexText "A"  `shouldBe` (Just $ CI 0)
      parseColumnIndexText "B"  `shouldBe` (Just $ CI 1)
      parseColumnIndexText "Z"  `shouldBe` (Just $ CI 25)
      parseColumnIndexText "AA" `shouldBe` (Just $ CI 26)
      parseColumnIndexText "AB" `shouldBe` (Just $ CI 27)
      parseColumnIndexText "BA" `shouldBe` (Just $ CI 52)
      parseColumnIndexText "AAA" `shouldBe` (Just $ CI 702)
      parseColumnIndexText "ABM" `shouldBe` (Just $ CI 740)
      parseColumnIndexText ""   `shouldBe` Nothing
      parseColumnIndexText "a"  `shouldBe` Nothing
    it "format column index" $ do
      formatColumnIndex (CI 0)   `shouldBe`  "A"
      formatColumnIndex (CI 1)   `shouldBe`  "B"
      formatColumnIndex (CI 25)  `shouldBe`  "Z"
      formatColumnIndex (CI 26)  `shouldBe`  "AA"
      formatColumnIndex (CI 27)  `shouldBe`  "AB"
      formatColumnIndex (CI 52)  `shouldBe`  "BA"
      formatColumnIndex (CI 702) `shouldBe`  "AAA"
      formatColumnIndex (CI 731) `shouldBe`  "ABD"
    it "parse cell index" $ do
      parseCellIndexText "A1"  `shouldBe` (Just $ CellIndex (RI 0) (CI 0))
      parseCellIndexText "AA2" `shouldBe` (Just $ CellIndex (RI 1) (CI 26))
    it "row movable" $ do
      (+) 1 <$.> "A1" `shouldBe` parseCellIndexTextUnsafe "A2"
      flip (-) 1 <$.> "A10" `shouldBe` parseCellIndexTextUnsafe "A9"
    it "column movable" $ do
      (+) 1 <$>> "A1" `shouldBe` parseCellIndexTextUnsafe "B1"
      flip (-) 1 <$>> "C1" `shouldBe` parseCellIndexTextUnsafe "B1"
