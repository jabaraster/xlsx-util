{-# LANGUAGE OverloadedStrings #-}
module Jabara.Util.XlsxSpec (spec) where

import           Jabara.Util.Xlsx
import           Test.Hspec
import           Test.Hspec.QuickCheck (prop)

spec :: Spec
spec = do
    describe "column label to index" $ do
        it "A to 1" $ do
            columnLabelToIndex "A" `shouldBe` 1
        it "Z to 26" $ do
            columnLabelToIndex "Z" `shouldBe` 26
        it "AA to 27" $ do
            columnLabelToIndex "AA" `shouldBe` 27
        it "AZ to 51" $ do
            columnLabelToIndex "AZ" `shouldBe` 52
    describe "default column name to index" $ do
        it "A1 to (1,1)" $ do
            parseCellName "A1" `shouldBe` (Just CellLocation {
                                                  _clRow    = 1
                                                , _clColumn = 1
                                                })
        it "AA1 to (1,27)" $ do
            parseCellName "AA1" `shouldBe` (Just CellLocation {
                                                   _clRow    = 1
                                                 , _clColumn = 27
                                                 })
        it "AZ23 to (23,52)" $ do
            parseCellName "AZ23" `shouldBe` (Just CellLocation {
                                                    _clRow    = 23
                                                  , _clColumn = 52
                                                  })

