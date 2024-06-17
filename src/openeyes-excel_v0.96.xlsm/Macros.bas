Attribute VB_Name = "Macros"
Option Explicit

' MarketDataInput ��ũ�� (Market Data sheet & vol sheet)
Sub MarketDataInputMacros()
    
    Call InputPrice.InputPrice
    
    Call InputYieldCurve.InputYieldCurve
        
    Call InputCorrelation.InputCorrelation
    
    Call Inputvol.Inputvol
     
End Sub

'MarketDataPost ��ũ�� (Market Data Sheet)
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

    Call
    
End Sub
