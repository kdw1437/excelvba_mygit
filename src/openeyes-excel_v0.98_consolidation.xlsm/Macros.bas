Attribute VB_Name = "Macros"
Option Explicit

' MarketDataInput 매크로 (Market Data sheet & vol sheet)
Sub MarketDataInputMacros()
    
    Call InputPrice.InputPrice
    
    Call InputYieldCurve.InputYieldCurve
        
    Call InputCorrelation.InputCorrelation
    
    Call Inputvol.Inputvol
     
End Sub

'MarketDataPost 매크로 (Market Data Sheet)
Sub MarketDataPostMacros()

    Call ClassPostPrice.ClassPostPrice
    
    Call ClassPostCorrhardcoded.PrintJsonString
    
    Call ClassPostCorrhardcoded.PrintJsonString2
    
    Call ClassPostYieldCurve.ClassPostYieldCurve
    
    Call ClassPostVol.RunFunc
    
End Sub

Sub VolDataInputMacro()
    
    Call Inputvol.Inputvol
    
End Sub

Sub VolDataPostMacro()

    Call ClassPostVol.RunFunc
        
End Sub


Sub DivStreamInputMacro()

    Call InputDivStream.InputDivStream
    
End Sub

Sub DivStreamPostMacro()

    Call ClassPostDivStream.ClassPostDivStream
        
End Sub

Sub DivYieldInputMacro()

    Call InputDivYield.InputDivYield
    
End Sub

Sub DivYieldPostMacro()

    Call ClassPostDivYield.ClassPostDivYield
    
End Sub

Sub QuotePost()

    Call ConvertRangeToJson.ConvertRangeToJson
        
End Sub

Sub Valuation()

    Call ValuationRequest.ValuationRequest
        
End Sub

Sub TestImport()
    ImportGreekData "81"
End Sub
