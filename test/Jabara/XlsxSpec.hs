{-# LANGUAGE OverloadedStrings #-}
module Jabara.XlsxSpec where

import           Codec.Xlsx
import           Control.Lens
import           Data.Either
import           Data.Maybe
import           Jabara.Xlsx
import           Test.Hspec

spec :: Spec
spec = do
  mBook <- runIO $ readBook "./kaji.xlsx"
  describe "readBook" $
    it "read book" $
      isRight mBook `shouldBe` True
  describe "cell values" $ do
    let book  = head $ rights [mBook]
        sheet = fromJust $ book ^? ixSheet "家計簿"
    it "text value" $
      cellStringValueFromSheet sheet (cellUnsafe "B1") `shouldBe` "家計簿"
    it "number value" $
      cellIntegralValueFromSheet sheet (cellUnsafe "C4") `shouldBe` (36000::Integer)
