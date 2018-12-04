{-# LANGUAGE OverloadedStrings #-}
module Jabara.Util.XlsxSpec where

import           Codec.Xlsx
import           Control.Lens
import           Data.Maybe
import           Jabara.Util.Xlsx
import           Test.Hspec

spec :: Spec
spec = do
  mBook <- runIO $ readBook "./kaji.xlsx"
  describe "readBook" $
    it "read book" $
      isJust mBook `shouldBe` True
  describe "cell values" $ do
    let book  = fromJust mBook
        sheet = fromJust $ book ^? ixSheet "家計簿"
    it "text value" $
      cellStringValueFromSheet sheet (CellLocation (0, 1)) `shouldBe` "家計簿"
    it "number value" $
      cellIntegralValueFromSheet sheet (3, 2) `shouldBe` (36000::Integer)
  describe "parse cell location text" $
    it "" $ do
      parseCellLocationText "A1" `shouldBe` (Just $ CellLocation (0, 0))
      parseCellLocationText "A10" `shouldBe` (Just $ CellLocation (9, 0))
      parseCellLocationText "B10" `shouldBe` (Just $ CellLocation (9, 1))
      parseCellLocationText "AB10" `shouldBe` (Just $ CellLocation (9, 27))
      parseCellLocationText "10" `shouldBe` Nothing
      parseCellLocationText "A" `shouldBe` Nothing
