Attribute VB_Name = "Module1"
Option Explicit

Enum header
    ファイル名 = 0
    シート名 = 1
    テキスト = 2

End Enum




Sub エクセル情報抽出()
 Dim Path As Variant
 Dim xl_wbk As Workbook
  
  Application.ScreenUpdating = False
  
  Set xl_wbk = Workbooks.Add '新規ワークブック追加
 
 
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
        With ws.ListObjects("ファイルパス")
        
        Dim i As Long
        Dim 抽出先セル
        
        
        For i = 1 To .ListRows.Count
            Path = .ListRows(i).Range.Value
            Set 抽出先セル = xl_wbk.Sheets("Sheet1").Cells(Rows.Count, 1).End(xlUp).Offset(1, 0)
        
            '-------------------------------------------------
            Call エクセル内の情報をExcelに集約出力する(Path, 抽出先セル)  '対象のマクロを選択
            '-------------------------------------------------
        
       Next i
    
    End With
  Application.ScreenUpdating = True
End Sub

Sub エクセル内の情報をExcelに集約出力する(Path, 抽出先セル)  '対象のマクロを選択

    Dim 抽出対象エクセル As Workbook
    
    'ファイルパスの参照先にファイルが存在しない場合、
    '元の処理に戻る
    If Dir(Path) = "" Then
        Exit Sub
    End If
    
    DoEvents
    Set 抽出対象エクセル = Workbooks.Open(Filename:=Path, UpdateLinks:=False)
    

    Dim シート
    
    
    For Each シート In 抽出対象エクセル.Sheets
    
        
        抽出先セル.Offset(0, header.ファイル名) = Path
        抽出先セル.Offset(0, header.シート名) = シート.Name
        
        
        Call シート内の情報を吐き出す(シート, 抽出先セル)
        'Debug.Print (シート.Name)
        'Debug.Print (抽出先セル)
    Next シート
    
    抽出対象エクセル.Close SaveChanges:=False
    

End Sub

Sub シート内の情報を吐き出す(抽出対象シート, 抽出先セル)

    Call オブジェクト内の情報を抽出する(抽出対象シート, 抽出先セル)
    Call セル内の情報を抽出する(抽出対象シート, 抽出先セル)
End Sub

Function オブジェクト内の情報を抽出する(抽出対象シート, 抽出先セル)

    Dim s_抽出対象シート As Worksheet
    
    Set s_抽出対象シート = 抽出対象シート
    'もし図形などが挿入されてなければ
    '次に進みます。
    If s_抽出対象シート.Shapes.Count >= 1 Then
        Dim 追加行 As Integer
        追加行 = 0
        Dim 図形 As Shape
        For Each 図形 In s_抽出対象シート.Shapes
            If 図形.TextFrame2.HasText Then
                
                
                抽出先セル.Offset(追加行, header.テキスト).Value = 図形.TextFrame2.TextRange.Text
                
              
                追加行 = 追加行 + 1
            
                抽出先セル.Offset(追加行, header.ファイル名).FormulaR1C1 = "=R[-1]C"
                抽出先セル.Offset(追加行, header.シート名).FormulaR1C1 = "=R[-1]C"
            
            End If
        Next 図形

        
    End If
    
    Set 抽出先セル = 抽出先セル.Offset(追加行, 0)
End Function


Function セル内の情報を抽出する(抽出対象シート, 抽出先セル)

    Dim s_抽出対象シート As Worksheet
    Set s_抽出対象シート = 抽出対象シート
    
    
    Dim 追加行 As Integer
    追加行 = 0
    
    Dim 調査範囲 As Range
    Dim セル As Range
    
    Set 調査範囲 = s_抽出対象シート.Range("A1").CurrentRegion
    For Each セル In 調査範囲
            If セル.Value <> "" Then
    
                抽出先セル.Offset(追加行, header.テキスト).Value = セル.Value
                
              
                追加行 = 追加行 + 1
            
                抽出先セル.Offset(追加行, header.ファイル名).FormulaR1C1 = "=R[-1]C"
                抽出先セル.Offset(追加行, header.シート名).FormulaR1C1 = "=R[-1]C"
    
            End If
    Next セル

                抽出先セル.Offset(追加行, header.ファイル名).Value = ""
                抽出先セル.Offset(追加行, header.シート名).Value = ""
End Function
