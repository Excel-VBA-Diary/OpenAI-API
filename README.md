# OpenAI-API
### OpenAI APIを経由して問い合わせた結果を返す. (Returns the results of a query made via the OpenAI API.)
  
Version 1.00 (September 14, 2026)
  
2026年9月14日以降、ExcelのCOPILOT関数が使用できなくなるというので代替関数としてGPT関数を作ってみた。  
そもそもCOPILOT関数がなくなるのはExcelの設計思想の一貫性を保つため。  
すなわち、Excelの設計思想（＝数式は“常に同じ入力なら同じ結果を返す”という決定性・再現性）と、AI応答の“非決定性”が根本的に噛み合わないためということらしい。  
したがってGPT関数もExcelの設計思想に合わない。  
それでも非決定性を容認しつつ利用したいケースはあるだろう。  
  
### 前提条件  
・コード内の OpenAI_API_KEY はご自身で取得したGemini APIキーを割り当てること。  
・Gemini APIキーは OpenAI Platform で取得する。  
  
### 使用例
```
Dim ans As String
ans = GPT("日本で一番高い山は？")
```
### 構文  
GPT((PromptText, \[ModelName\])  
  
GPT関数の書式には、次の引数があります。   
|  引数  |  説明  |
| ------ | ------ |
| PromptText（必須） | プロンプト |  
| ModelNam（省略可） | AIモデルを指定する。デフォルトはgpt-5.6-luna 。 |  
  
OpenAIの利用ライセンス（料金プラン）によって利用できるAIモデルや回数制限が異なるので事前に確認すること。  
  
### VBAコード
```
Option Explicit

'-------------------------------------------------------------------------------
' OpenAI APIを経由して問い合わせた結果を返す.
' Returns the results of a query made via the OpenAI API.
'
' Syntax:
'   GPT(PromptText, [ModelName])
'
' Arguments:
'   PromptText: Required, String
'               Text that describes the task or question for the AI model.
'               PromptText must be a single prompt.
'
'   ModelName: Optional, String
'              Specify the OpenAI AI model. If omitted,
'              the default is "gpt-5.6-luna.
'
' Return Value:
'   Response (String)
'
' Usage:
'   Dim ans As String
'   ans = GPT(prompt)
'
' 前提条件
'   コード内の OpenAI_API_KEY はご自身で取得したOpenAI APIキーを割り当てること。
'   APIキーは OpenAI Platform で取得する。
'
' Prerequisites
'   Replace “OpenAI_API_KEY” in the code with the OpenAI API key you obtained.
'   You can obtain the API key from OpenAI Platform.
'
'
' Author: Excel VBA Diary (@excelvba_diary)
' Created: September 14, 2026
' Last Updated: September 14, 2026
' Version: 1.000
' License: MIT
'-------------------------------------------------------------------------------


Private Const DefaultModel As String = "gpt-5.6-luna"

Public Function GPT(ByVal PromptText As String, _
                    Optional ModelName As String = "") As String

    On Error GoTo ErrHandler
    
    If Trim(PromptText) = "" Then
        GPT = "#ERROR: EmptyPrompt"
        Exit Function
    End If
    
    If Len(OpenAI_API_KEY) = 0 Then
        GPT = "#ERROR: NoApiKey"
        Exit Function
    End If
    
    Dim model As String
    model = IIf(ModelName = "", DefaultModel, ModelName)

    ' プロンプト内の特殊文字（\, ", CRLF）をJSON用にエスケープする
    ' Escape special characters (\, ", line breaks) in the prompt for JSON
    
    Dim safePrompt As String
    safePrompt = PromptText
    safePrompt = Replace(safePrompt, "\", "\\")
    safePrompt = Replace(safePrompt, """", "\""")
    safePrompt = Replace(safePrompt, vbCrLf, "\n")
    safePrompt = Replace(safePrompt, vbCr, "\n")
    safePrompt = Replace(safePrompt, vbLf, "\n")
    
    ' JSONペイロードの作成
    ' Creating a JSON Payload
    
    Dim jsonPpayload As String
    jsonPpayload = _
    "{""model"":""" & model & """," & _
        """input"":[{" & _
            """role"":""user""," & _
            """content"":[{" & _
                """type"":""input_text""," & _
                """text"":""" & PromptText & """" & _
                "}]" & _
            "}]," & _
            """text"":{" & _
                """format"":{" & _
                    """type"":""json_schema""," & _
                    """name"":""answer_schema""," & _
                    """strict"":true," & _
                    """schema"":{" & _
                        """type"":""object""," & _
                        """properties"":{""answer"":{""type"":""string""}}," & _
                        """required"":[""answer""],""additionalProperties"":false" & _
                    "}" & _
                "}" & _
            "}" & _
        "}"

    Dim api_url As String
    api_url = "https://api.openai.com/v1/responses"
    
    Dim objHttp As Object
    Set objHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    With objHttp
        .Open "POST", api_url, False
        .setTimeouts 5000, 5000, 10000, 60000
        .setRequestHeader "Content-Type", "application/json; charset=utf-8"
        .setRequestHeader "Authorization", "Bearer " & OpenAI_API_KEY
        .send StrToUtf8Bytes(jsonPpayload)
        
        If .Status <> 200 Then
            GPT = "#ERROR:" & .Status & ":" & .responseText
            Debug.Print "POST Error occurred: Status="; .Status
            Debug.Print .responseText
            Exit Function
        End If
            
        Dim jsonResponse As String
        jsonResponse = DecodeUtf8Response(.responseBody)
    End With
    
    Dim keyValue As String
    keyValue = ExtractJsonValue(jsonResponse, "answer")
    keyValue = UnescapeUnicode(keyValue)

    GPT = keyValue

    Exit Function

ErrHandler:
    GPT = "#ERROR:Exception:" & Err.Number & ":" & Err.Description
    Debug.Print "Runtime Error occurred: Number="; Err.Number
    Debug.Print Err.Description
    
End Function


' JSONのテキストデーターをバイナリーデーターに変換する
' Convert JSON text data to binary data

Private Function StrToUtf8Bytes(ByVal JsonText As String) As Variant
    
    Dim objStream As Object
    Set objStream = CreateObject("ADODB.Stream")
    
    With objStream
        .Type = 2                   ' adTypeText (Text Data)
        .Charset = "UTF-8"
        .Open
        .WriteText JsonText
        .Position = 0
        .Type = 1                   ' adTypeBinary (Binary Data)
        .Position = 3               ' Since a UTF-8 BOM (EF BB BF) is appended at the beginning, skip 3 bytes.
        StrToUtf8Bytes = .Read      ' An array is returned as a Variant
        .Close
    End With

End Function


' JSONのバイナリーデータをテキストデーターに変換する
' Convert JSON binary data to text data

Private Function DecodeUtf8Response(ByVal JsonBinary As Variant) As String
    
    Dim objStream As Object
    Set objStream = CreateObject("ADODB.Stream")
    
    With objStream
        .Type = 1                   ' adTypeBinary (Binary Data)
        .Open
        .Write JsonBinary
        .Position = 0
        .Type = 2                   ' adTypeText (Text Data)
        .Charset = "UTF-8"
        DecodeUtf8Response = .ReadText
        .Close
    End With

End Function


' JSONから指定したキーの文字列を取得する
' ここでは正規表現で抽出しているがJSONパーサーを使ってもよい

' Retrieve the string for a specified key from JSON
' Although this code is using regular expressions for extraction here,
' you can also use a JSON parser.

Function ExtractJsonValue(ByVal JsonText As String, _
                          ByVal keyName As String) As String

    Dim matches As Object, strTemp As String
    
    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    
    With objRegExp
        .Pattern = """text"":\s*?""{\\""answer\\"":\\""(.*?)\\""}"
        Set matches = .Execute(JsonText)
        If matches.Count > 0 Then
            strTemp = matches(0).SubMatches(0)
            strTemp = Replace(strTemp, "\""", """")
            strTemp = Replace(strTemp, "\\", "\")
            strTemp = Replace(strTemp, "\n", vbCrLf)
            ExtractJsonValue = strTemp
        Else
            ExtractJsonValue = "#ERROR: TextNotFound"
        End If
    End With

End Function


' Unicodeエスケープ形式（\uxxxx）の文字列をデコードする
' Decode a string in Unicode escape format (\uxxxx)

Private Function UnescapeUnicode(ByVal EscapedText As String) As String

    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    
    With objRegExp
        .Pattern = "\\u([0-9a-fA-F]{4})"
        .IgnoreCase = True
        .Global = True
        Dim matches As Object
        Set matches = .Execute(EscapedText)
    End With

    If matches.Count = 0 Then
        UnescapeUnicode = ""
        Exit Function
    End If

    ' マッチした \uXXXX を実際の文字に置き換える
    ' Replace the matched \uXXXX with the actual character
    
    Dim decodedText As String, hexCode As String, char As String
    decodedText = EscapedText

    Dim match As Object
    For Each match In matches
        hexCode = match.SubMatches(0)
        char = ChrW(CLng("&H" + hexCode))
        decodedText = Replace(decodedText, match.Value, char)
    Next match
    
    UnescapeUnicode = decodedText
End Function
```
  

