param(
    [string]$Path,
    [switch]$Uninstall,
    # Tvinger reinstallasjon selv om versjonen er lik (brukes av EMI "Reinstaller alle").
    [switch]$Force,
    # Kun for testing: start en EGEN Excel-prosess i stedet for a koble til den
    # som allerede kjorer. Samme monster som Kontaktsentralen.ps1.
    [switch]$NyExcelInstans
)

$ErrorActionPreference = 'Stop'

# ============================================================================
# Nettskjema Henter - alt-i-en installer/oppdaterer.
# Bygger Ã¥tte VBA-komponenter fra kildekoden innebygd nedenfor (JsonConverter,
# modCredentialManager, modAuthentication, modJsonFlatten, modNettskjema,
# frmNettskjemaKoblinger, frmNettskjemaInnstillinger, frmKolonneInnstillinger)
# i en midlertidig arbeidsbok, og installerer/
# oppdaterer dem deretter i målfilen. Ingen andre filer trengs eller
# etterlates.
#
# Bruk: kjor mot en eksisterende .xlsm mens arket du vil ha knappen på er
# aktivt (samme mateer som Kolonnevelger/Mail-utsender). Etter installasjon:
# trykk knappen "Nettskjema Henter" og bruk "+ Ny tilkobling..." for å
# koble ett eller flere Nettskjema-skjemaer - hvert skjema får sitt eget
# ark + tabell ("<arknavn>_tbl"). Et skjult register i arbeidsboken husker
# hvilke Form ID-er som allerede er koblet til hvilket ark.
# ============================================================================

# ---- Tredjeparts JSON-parser (VBA-JSON v2.3.1, MIT-lisens) ----
# https://github.com/VBA-tools/VBA-JSON - IKKE endre denne utover den ene
# tilpasningen (Dictionary -> Scripting.Dictionary, sen-bundet, for å unngå
# en VBA-prosjektreferanse som kan mangle på andre maskiner - samme mateer
# som resten av TIMSS-verktoyene allerede bruker for ordboker).

$jsonConverterCode = @'
Attribute VB_Name = "JsonConverter"
''
' JSON Converter for VBA - basert på VBA-JSON v2.3.1 (Tim Hall, MIT-lisens)
' https://github.com/VBA-tools/VBA-JSON
'
' TILPASSET for dette prosjektet: Dictionary-objekter opprettes sen-bundet
' via CreateObject("Scripting.Dictionary") i stedet for tidlig-bundet
' "New Dictionary", slik at ingen ekstra VBA-prosjektreferanse (Microsoft
' Scripting Runtime) trengs. Resten av koden er uendret.
''
Option Explicit

Private Type json_Options
    UseDoubleForLargeNumbers As Boolean
    AllowUnquotedKeys As Boolean
    EscapeSolidus As Boolean
End Type
Public JsonOptions As json_Options

' ============================================= '
' Public Methods
' ============================================= '

Public Function ParseJson(ByVal JsonString As String) As Object
    Dim json_Index As Long
    json_Index = 1

    JsonString = VBA.Replace(VBA.Replace(VBA.Replace(JsonString, VBA.vbCr, ""), VBA.vbLf, ""), VBA.vbTab, "")

    json_SkipSpaces JsonString, json_Index
    Select Case VBA.Mid$(JsonString, json_Index, 1)
    Case "{"
        Set ParseJson = json_ParseObject(JsonString, json_Index)
    Case "["
        Set ParseJson = json_ParseArray(JsonString, json_Index)
    Case Else
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(JsonString, json_Index, "Expecting '{' or '['")
    End Select
End Function

Public Function ConvertToJson(ByVal JsonValue As Variant, Optional ByVal Whitespace As Variant, Optional ByVal json_CurrentIndentation As Long = 0) As String
    Dim json_Buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long
    Dim json_Index As Long
    Dim json_LBound As Long
    Dim json_UBound As Long
    Dim json_IsFirstItem As Boolean
    Dim json_Key As Variant
    Dim json_Value As Variant
    Dim json_DateStr As String
    Dim json_Converted As String
    Dim json_SkipItem As Boolean
    Dim json_PrettyPrint As Boolean
    Dim json_Indentation As String

    json_LBound = -1
    json_UBound = -1
    json_IsFirstItem = True
    json_PrettyPrint = Not IsMissing(Whitespace)

    Select Case VBA.VarType(JsonValue)
    Case VBA.vbNull
        ConvertToJson = "null"
    Case VBA.vbDate
        json_DateStr = ConvertToIso(VBA.CDate(JsonValue))
        ConvertToJson = """" & json_DateStr & """"
    Case VBA.vbString
        If Not JsonOptions.UseDoubleForLargeNumbers And json_StringIsLargeNumber(JsonValue) Then
            ConvertToJson = JsonValue
        Else
            ConvertToJson = """" & json_Encode(JsonValue) & """"
        End If
    Case VBA.vbBoolean
        If JsonValue Then
            ConvertToJson = "true"
        Else
            ConvertToJson = "false"
        End If
    Case VBA.vbObject
        If json_PrettyPrint Then
            json_Indentation = VBA.Space$((json_CurrentIndentation + 1) * Whitespace)
        End If

        If VBA.TypeName(JsonValue) = "Dictionary" Then
            json_BufferAppend json_Buffer, "{", json_BufferPosition, json_BufferLength
            For Each json_Key In JsonValue.Keys
                json_Converted = ConvertToJson(JsonValue(json_Key), Whitespace, json_CurrentIndentation + 1)
                If json_Converted = "" Then
                    json_SkipItem = json_IsUndefined(JsonValue(json_Key))
                Else
                    json_SkipItem = False
                End If

                If Not json_SkipItem Then
                    If json_IsFirstItem Then
                        json_IsFirstItem = False
                    Else
                        json_BufferAppend json_Buffer, ",", json_BufferPosition, json_BufferLength
                    End If

                    If json_PrettyPrint Then
                        json_Converted = vbNewLine & json_Indentation & """" & json_Key & """: " & json_Converted
                    Else
                        json_Converted = """" & json_Key & """:" & json_Converted
                    End If

                    json_BufferAppend json_Buffer, json_Converted, json_BufferPosition, json_BufferLength
                End If
            Next json_Key

            If json_PrettyPrint Then
                json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength
                json_Indentation = VBA.Space$(json_CurrentIndentation * Whitespace)
            End If

            json_BufferAppend json_Buffer, json_Indentation & "}", json_BufferPosition, json_BufferLength

        ElseIf VBA.TypeName(JsonValue) = "Collection" Then
            json_BufferAppend json_Buffer, "[", json_BufferPosition, json_BufferLength
            For Each json_Value In JsonValue
                If json_IsFirstItem Then
                    json_IsFirstItem = False
                Else
                    json_BufferAppend json_Buffer, ",", json_BufferPosition, json_BufferLength
                End If

                json_Converted = ConvertToJson(json_Value, Whitespace, json_CurrentIndentation + 1)
                If json_Converted = "" And json_IsUndefined(json_Value) Then json_Converted = "null"

                If json_PrettyPrint Then
                    json_Converted = vbNewLine & json_Indentation & json_Converted
                End If

                json_BufferAppend json_Buffer, json_Converted, json_BufferPosition, json_BufferLength
            Next json_Value

            If json_PrettyPrint Then
                json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength
                json_Indentation = VBA.Space$(json_CurrentIndentation * Whitespace)
            End If

            json_BufferAppend json_Buffer, json_Indentation & "]", json_BufferPosition, json_BufferLength
        End If

        ConvertToJson = json_BufferToString(json_Buffer, json_BufferPosition)
    Case VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal
        ConvertToJson = VBA.Replace(JsonValue, ",", ".")
    Case Else
        On Error Resume Next
        ConvertToJson = JsonValue
        On Error GoTo 0
    End Select
End Function

' ============================================= '
' Private Functions
' ============================================= '

Private Function json_ParseObject(json_String As String, ByRef json_Index As Long) As Object
    Dim json_Key As String
    Dim json_NextChar As String

    Set json_ParseObject = CreateObject("Scripting.Dictionary")
    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> "{" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '{'")
    Else
        json_Index = json_Index + 1

        Do
            json_SkipSpaces json_String, json_Index
            If VBA.Mid$(json_String, json_Index, 1) = "}" Then
                json_Index = json_Index + 1
                Exit Function
            ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
                json_Index = json_Index + 1
                json_SkipSpaces json_String, json_Index
            End If

            json_Key = json_ParseKey(json_String, json_Index)
            json_NextChar = json_Peek(json_String, json_Index)
            If json_NextChar = "[" Or json_NextChar = "{" Then
                Set json_ParseObject.Item(json_Key) = json_ParseValue(json_String, json_Index)
            Else
                json_ParseObject.Item(json_Key) = json_ParseValue(json_String, json_Index)
            End If
        Loop
    End If
End Function

Private Function json_ParseArray(json_String As String, ByRef json_Index As Long) As Collection
    Set json_ParseArray = New Collection

    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> "[" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '['")
    Else
        json_Index = json_Index + 1

        Do
            json_SkipSpaces json_String, json_Index
            If VBA.Mid$(json_String, json_Index, 1) = "]" Then
                json_Index = json_Index + 1
                Exit Function
            ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
                json_Index = json_Index + 1
                json_SkipSpaces json_String, json_Index
            End If

            json_ParseArray.Add json_ParseValue(json_String, json_Index)
        Loop
    End If
End Function

Private Function json_ParseValue(json_String As String, ByRef json_Index As Long) As Variant
    json_SkipSpaces json_String, json_Index
    Select Case VBA.Mid$(json_String, json_Index, 1)
    Case "{"
        Set json_ParseValue = json_ParseObject(json_String, json_Index)
    Case "["
        Set json_ParseValue = json_ParseArray(json_String, json_Index)
    Case """", "'"
        json_ParseValue = json_ParseString(json_String, json_Index)
    Case Else
        If VBA.Mid$(json_String, json_Index, 4) = "true" Then
            json_ParseValue = True
            json_Index = json_Index + 4
        ElseIf VBA.Mid$(json_String, json_Index, 5) = "false" Then
            json_ParseValue = False
            json_Index = json_Index + 5
        ElseIf VBA.Mid$(json_String, json_Index, 4) = "null" Then
            json_ParseValue = Null
            json_Index = json_Index + 4
        ElseIf VBA.InStr("+-0123456789", VBA.Mid$(json_String, json_Index, 1)) Then
            json_ParseValue = json_ParseNumber(json_String, json_Index)
        Else
            Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['")
        End If
    End Select
End Function

Private Function json_ParseString(json_String As String, ByRef json_Index As Long) As String
    Dim json_Quote As String
    Dim json_Char As String
    Dim json_Code As String
    Dim json_Buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long

    json_SkipSpaces json_String, json_Index

    json_Quote = VBA.Mid$(json_String, json_Index, 1)
    json_Index = json_Index + 1

    Do While json_Index > 0 And json_Index <= Len(json_String)
        json_Char = VBA.Mid$(json_String, json_Index, 1)

        Select Case json_Char
        Case "\"
            json_Index = json_Index + 1
            json_Char = VBA.Mid$(json_String, json_Index, 1)

            Select Case json_Char
            Case """", "\", "/", "'"
                json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "b"
                json_BufferAppend json_Buffer, vbBack, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "f"
                json_BufferAppend json_Buffer, vbFormFeed, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "n"
                json_BufferAppend json_Buffer, vbCrLf, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "r"
                json_BufferAppend json_Buffer, vbCr, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "t"
                json_BufferAppend json_Buffer, vbTab, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "u"
                json_Index = json_Index + 1
                json_Code = VBA.Mid$(json_String, json_Index, 4)
                json_BufferAppend json_Buffer, VBA.ChrW(VBA.Val("&h" + json_Code)), json_BufferPosition, json_BufferLength
                json_Index = json_Index + 4
            End Select
        Case json_Quote
            json_ParseString = json_BufferToString(json_Buffer, json_BufferPosition)
            json_Index = json_Index + 1
            Exit Function
        Case Else
            json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
            json_Index = json_Index + 1
        End Select
    Loop
End Function

Private Function json_ParseNumber(json_String As String, ByRef json_Index As Long) As Variant
    Dim json_Char As String
    Dim json_Value As String
    Dim json_IsLargeNumber As Boolean

    json_SkipSpaces json_String, json_Index

    Do While json_Index > 0 And json_Index <= Len(json_String)
        json_Char = VBA.Mid$(json_String, json_Index, 1)

        If VBA.InStr("+-0123456789.eE", json_Char) Then
            json_Value = json_Value & json_Char
            json_Index = json_Index + 1
        Else
            json_IsLargeNumber = IIf(InStr(json_Value, "."), Len(json_Value) >= 17, Len(json_Value) >= 16)
            If Not JsonOptions.UseDoubleForLargeNumbers And json_IsLargeNumber Then
                json_ParseNumber = json_Value
            Else
                json_ParseNumber = VBA.Val(json_Value)
            End If
            Exit Function
        End If
    Loop
End Function

Private Function json_ParseKey(json_String As String, ByRef json_Index As Long) As String
    If VBA.Mid$(json_String, json_Index, 1) = """" Or VBA.Mid$(json_String, json_Index, 1) = "'" Then
        json_ParseKey = json_ParseString(json_String, json_Index)
    ElseIf JsonOptions.AllowUnquotedKeys Then
        Dim json_Char As String
        Do While json_Index > 0 And json_Index <= Len(json_String)
            json_Char = VBA.Mid$(json_String, json_Index, 1)
            If (json_Char <> " ") And (json_Char <> ":") Then
                json_ParseKey = json_ParseKey & json_Char
                json_Index = json_Index + 1
            Else
                Exit Do
            End If
        Loop
    Else
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '""' or '''")
    End If

    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> ":" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting ':'")
    Else
        json_Index = json_Index + 1
    End If
End Function

Private Function json_IsUndefined(ByVal json_Value As Variant) As Boolean
    Select Case VBA.VarType(json_Value)
    Case VBA.vbEmpty
        json_IsUndefined = True
    Case VBA.vbObject
        Select Case VBA.TypeName(json_Value)
        Case "Empty", "Nothing"
            json_IsUndefined = True
        End Select
    End Select
End Function

Private Function json_Encode(ByVal json_Text As Variant) As String
    Dim json_Index As Long
    Dim json_Char As String
    Dim json_AscCode As Long
    Dim json_Buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long

    For json_Index = 1 To VBA.Len(json_Text)
        json_Char = VBA.Mid$(json_Text, json_Index, 1)
        json_AscCode = VBA.AscW(json_Char)

        If json_AscCode < 0 Then
            json_AscCode = json_AscCode + 65536
        End If

        Select Case json_AscCode
        Case 34
            json_Char = "\"""
        Case 92
            json_Char = "\\"
        Case 47
            If JsonOptions.EscapeSolidus Then json_Char = "\/"
        Case 8
            json_Char = "\b"
        Case 12
            json_Char = "\f"
        Case 10
            json_Char = "\n"
        Case 13
            json_Char = "\r"
        Case 9
            json_Char = "\t"
        Case 0 To 31, 127 To 65535
            json_Char = "\u" & VBA.Right$("0000" & VBA.Hex$(json_AscCode), 4)
        End Select

        json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
    Next json_Index

    json_Encode = json_BufferToString(json_Buffer, json_BufferPosition)
End Function

Private Function json_Peek(json_String As String, ByVal json_Index As Long, Optional json_NumberOfCharacters As Long = 1) As String
    json_SkipSpaces json_String, json_Index
    json_Peek = VBA.Mid$(json_String, json_Index, json_NumberOfCharacters)
End Function

Private Sub json_SkipSpaces(json_String As String, ByRef json_Index As Long)
    Do While json_Index > 0 And json_Index <= VBA.Len(json_String) And VBA.Mid$(json_String, json_Index, 1) = " "
        json_Index = json_Index + 1
    Loop
End Sub

Private Function json_StringIsLargeNumber(json_String As Variant) As Boolean
    Dim json_Length As Long
    Dim json_CharIndex As Long
    json_Length = VBA.Len(json_String)

    If json_Length >= 16 And json_Length <= 100 Then
        Dim json_CharCode As String

        json_StringIsLargeNumber = True

        For json_CharIndex = 1 To json_Length
            json_CharCode = VBA.Asc(VBA.Mid$(json_String, json_CharIndex, 1))
            Select Case json_CharCode
            Case 46, 48 To 57, 69, 101
            Case Else
                json_StringIsLargeNumber = False
                Exit Function
            End Select
        Next json_CharIndex
    End If
End Function

Private Function json_ParseErrorMessage(json_String As String, ByRef json_Index As Long, ErrorMessage As String)
    Dim json_StartIndex As Long
    Dim json_StopIndex As Long

    json_StartIndex = json_Index - 10
    json_StopIndex = json_Index + 10
    If json_StartIndex <= 0 Then json_StartIndex = 1
    If json_StopIndex > VBA.Len(json_String) Then json_StopIndex = VBA.Len(json_String)

    json_ParseErrorMessage = "Error parsing JSON:" & VBA.vbNewLine & _
                             VBA.Mid$(json_String, json_StartIndex, json_StopIndex - json_StartIndex + 1) & VBA.vbNewLine & _
                             VBA.Space$(json_Index - json_StartIndex) & "^" & VBA.vbNewLine & _
                             ErrorMessage
End Function

Private Sub json_BufferAppend(ByRef json_Buffer As String, ByRef json_Append As Variant, ByRef json_BufferPosition As Long, ByRef json_BufferLength As Long)
    Dim json_AppendLength As Long
    Dim json_LengthPlusPosition As Long

    json_AppendLength = VBA.Len(json_Append)
    json_LengthPlusPosition = json_AppendLength + json_BufferPosition

    If json_LengthPlusPosition > json_BufferLength Then
        Dim json_AddedLength As Long
        json_AddedLength = IIf(json_AppendLength > json_BufferLength, json_AppendLength, json_BufferLength)

        json_Buffer = json_Buffer & VBA.Space$(json_AddedLength)
        json_BufferLength = json_BufferLength + json_AddedLength
    End If

    Mid$(json_Buffer, json_BufferPosition + 1, json_AppendLength) = CStr(json_Append)
    json_BufferPosition = json_BufferPosition + json_AppendLength
End Sub

Private Function json_BufferToString(ByRef json_Buffer As String, ByVal json_BufferPosition As Long) As String
    If json_BufferPosition > 0 Then
        json_BufferToString = VBA.Left$(json_Buffer, json_BufferPosition)
    End If
End Function

' ---- Minimal dato-hjelper (kun det denne løsningen trenger - ingen full VBA-UTC) ----

Public Function ConvertToIso(ByVal LocalDate As Date) As String
    ConvertToIso = VBA.Format$(LocalDate, "yyyy-mm-ddTHH:mm:ss")
End Function
'@

# ---- modCredentialManager: Windows Credential Manager-wrapper ----
# Fra og med 3.20.0 er arbeidsbokens VeryHidden registerark PRIMÆR lagring av
# Client Secret (se GetGlobalClientSecret/SetGlobalClientSecret i modNettskjema),
# slik at samme .xlsm kan brukes av alle autoriserte brukere. Denne modulen
# brukes nå kun som migreringsfallback fra eldre versjoner (en Client Secret
# lagret lokalt før 3.20.0) - se OfferClientSecretMigrationIfNeeded.

$credMgrCode = @'
Attribute VB_Name = "modCredentialManager"
Option Explicit

' Wrapper rundt Windows Credential Manager (advapi32.dll) for å lagre/lese
' Client Secret uten å noensinne skrive den i klartekst i arbeidsboken.

Private Const CRED_TYPE_GENERIC As Long = 1
Private Const CRED_PERSIST_LOCAL_MACHINE As Long = 2

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type CREDENTIAL
    Flags As Long
    CredType As Long
    TargetName As LongPtr
    Comment As LongPtr
    LastWritten As FILETIME
    CredentialBlobSize As Long
    CredentialBlob As LongPtr
    Persist As Long
    AttributeCount As Long
    Attributes As LongPtr
    TargetAlias As LongPtr
    UserName As LongPtr
End Type

Private Declare PtrSafe Function CredWriteW Lib "advapi32.dll" (ByRef Credential As CREDENTIAL, ByVal Flags As Long) As Long
Private Declare PtrSafe Function CredReadW Lib "advapi32.dll" (ByVal TargetName As LongPtr, ByVal CredType As Long, ByVal Flags As Long, ByRef CredentialPtr As LongPtr) As Long
Private Declare PtrSafe Function CredDeleteW Lib "advapi32.dll" (ByVal TargetName As LongPtr, ByVal CredType As Long, ByVal Flags As Long) As Long
Private Declare PtrSafe Sub CredFree Lib "advapi32.dll" (ByVal Buffer As LongPtr)
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' Lagrer (overskriver) et passord under et gitt målnavn.
Public Function CredMgr_Save(ByVal TargetName As String, ByVal UserName As String, ByVal Secret As String) As Boolean
    Dim cred As CREDENTIAL
    Dim tName As String, uName As String, blob As String
    Dim ok As Long

    tName = TargetName
    uName = UserName
    blob = Secret

    With cred
        .Flags = 0
        .CredType = CRED_TYPE_GENERIC
        .TargetName = StrPtr(tName)
        .Comment = 0
        .CredentialBlobSize = LenB(blob)
        .CredentialBlob = StrPtr(blob)
        .Persist = CRED_PERSIST_LOCAL_MACHINE
        .AttributeCount = 0
        .Attributes = 0
        .TargetAlias = 0
        .UserName = StrPtr(uName)
    End With

    ok = CredWriteW(cred, 0)
    CredMgr_Save = (ok <> 0)
End Function

' Leser et tidligere lagret passord. Returnerer "" hvis det ikke finnes.
Public Function CredMgr_Read(ByVal TargetName As String) As String
    Dim tName As String
    Dim pCred As LongPtr
    Dim cred As CREDENTIAL
    Dim ok As Long

    tName = TargetName
    pCred = 0
    ok = CredReadW(StrPtr(tName), CRED_TYPE_GENERIC, 0, pCred)
    If ok = 0 Or pCred = 0 Then
        CredMgr_Read = ""
        Exit Function
    End If

    CopyMemory cred, ByVal pCred, LenB(cred)
    CredMgr_Read = json_PtrToUnicodeString(cred.CredentialBlob, cred.CredentialBlobSize)
    CredFree pCred
End Function

' Sletter et lagret passord. Returnerer True selv om det ikke fantes fra før.
Public Function CredMgr_Delete(ByVal TargetName As String) As Boolean
    Dim tName As String
    tName = TargetName
    CredMgr_Delete = (CredDeleteW(StrPtr(tName), CRED_TYPE_GENERIC, 0) <> 0)
End Function

Private Function json_PtrToUnicodeString(ByVal Ptr As LongPtr, ByVal ByteLen As Long) As String
    Dim s As String
    If ByteLen <= 0 Or Ptr = 0 Then Exit Function
    s = String$(ByteLen \ 2, vbNullChar)
    CopyMemory ByVal StrPtr(s), ByVal Ptr, ByteLen
    json_PtrToUnicodeString = s
End Function
'@

# ---- modAuthentication: OAuth client-credentials-flyt + HTTP GET med retry ----

$authCode = @'
Attribute VB_Name = "modAuthentication"
Option Explicit

Public Const TOKEN_URL As String = "https://authorization.nettskjema.no/oauth2/token"
Public Const API_BASE_URL As String = "https://nettskjema.no/api/v3"
Public Const MAX_RETRIES As Long = 5
Public Const MAX_BACKOFF_SECONDS As Long = 30
Public Const TOKEN_EXPIRY_SAFETY_SECONDS As Long = 60
Public Const USER_AGENT As String = "Excel-Nettskjema-VBA/1.0"
Public Const CRED_TARGET_PREFIX As String = "NettskjemaHenter:"

' Satt til True hvis et Ctrl+Break/Esc-avbrudd blir fanget MELLOM to HTTP-
' forsøk (se sjekkene nederst i retry-lokkene under) - IKKE en garanti om at
' et enkelt, pågående kall kan avbrytes midt i (se historikk-notatet under
' om hvorfor det ble forsøkt og reversert). Holdt som Public/harmløs stubb
' av bakoverkompatibilitet med Håkons egen OppdaterNettskjema-makro, som
' kaller NettskjemaNullstillAvbrudd/ErHentingAvbrutt.
Public NettskjemaAvbrytHenting As Boolean

' Token caches KUN i minnet for denne Excel-økten - lagres aldri på disk.
' En arbeidsbok kan ha flere tilkoblinger med ULIKE API-klienter, så
' tokenet caches per ClientId (ikke i en enkelt modulvariabel).
Private mTokens As Object          ' ClientId -> access token
Private mTokenExpiries As Object   ' ClientId -> utløpstidspunkt (lokal tid)

Public Sub NettskjemaNullstillAvbrudd()
    NettskjemaAvbrytHenting = False
End Sub

Public Function ErHentingAvbrutt() As Boolean
    ErHentingAvbrutt = NettskjemaAvbrytHenting
End Function

' HISTORIKK (2026-09-28): 3.23.0-3.23.3 gjorde alle HTTP-kall asynkrone
' (http.Open ..., True) + en VentPaaHttpSvar-hjelper som polled
' WaitForResponse i 1s-bolker, for å gjøre en fastlåst henting avbrytbar
' med Esc i stedet for å måtte tvangslukke Excel. Esc-avbruddet ble
' bekreftet fungerende av Håkon selv - MEN etter det rapporterte han at
' selve datahentingen sluttet å fungere i det hele tatt («før kom dataen
' relativt raskt, men nå kommer den ikke i det hele tatt»). Mistanke (ikke
' endelig bekreftet, kunne ikke testes trygt mot hans ekte API): et stort,
' chunked/strømmende application/x-ndjson-svar for /answers oppfører seg
' annerledes mot ekte nettskjema.no i asynkron WinHttp-modus enn mot en
' liten, lokal test-respons - «WaitForResponse» ga aldri True, og kallet
' endte alltid i et tidsavbrudd etter det harde taket, uansett hvor lenge
' man ventet. Reversert tilbake til SYNKRON Send (slik det fungerte "før")
' på Håkons eksplisitte instruks, siden pålitelig datahenting er
' viktigere enn avbrytbarhet. Se g-nettskjema-http-synkron-blokkering for
' full detalj og en åpen idé til en TRYGGERE vei tilbake til avbrytbarhet
' senere (kjøre selve HTTP-kallet i en egen, ekte bakgrunnsprosess i
' stedet for asynkron WinHttp inne i VBA-tråden).

Private Sub EnsureTokenCache()
    If mTokens Is Nothing Then Set mTokens = CreateObject("Scripting.Dictionary")
    If mTokenExpiries Is Nothing Then Set mTokenExpiries = CreateObject("Scripting.Dictionary")
End Sub

' Tving ny token neste gang for denne klienten (brukes når en ny Client
' Secret lagres, og etter HTTP 401).
Public Sub ForgetAccessToken(ByVal ClientId As String)
    EnsureTokenCache
    If mTokens.Exists(ClientId) Then mTokens.Remove ClientId
    If mTokenExpiries.Exists(ClientId) Then mTokenExpiries.Remove ClientId
End Sub

Private Sub EnsureAccessToken(ByVal ClientId As String)
    EnsureTokenCache
    If Not mTokens.Exists(ClientId) Then
        AcquireAccessToken ClientId
    ElseIf Now >= mTokenExpiries(ClientId) Then
        AcquireAccessToken ClientId
    End If
End Sub

' BEKREFTET VIA DRIFTSLOGGEN (2026-09-16): et par forbigående, uforklarte
' nettverksfeil mot AKKURAT authorization.nettskjema.no (ikke selve
' dataendepunktet nettskjema.no/api/v3, som har sin egen retry i
' ApiGetWithRetry) endte opp som en RÅ, ufanget COM-feil ("Operasjonen ble
' tidsavbrutt") helt uten forsøk på ny - fordi denne funksjonen tidligere
' bare prøvde ÉN gang og ikke hadde noen retry/backoff i det hele tatt, i
' motsetning til /answers-kallet. Samme retry-mønster (MAX_RETRIES forsøk,
' eksponentiell backoff) som ApiGetWithRetry er derfor lagt til her også.
Private Sub AcquireAccessToken(ByVal ClientId As String)
    Dim clientSecret As String
    Dim json As Object
    Dim expiresIn As Long
    Dim attempt As Long
    Dim http As Object
    Dim sendFailed As Boolean
    Dim sendErrDescription As String
    Dim lastErrMsg As String

    ' Client Secret følger arbeidsboken slik at samme .xlsm kan brukes av alle
    ' autoriserte brukere uten lokalt Credential Manager-oppsett. VeryHidden
    ' er IKKE en sikkerhetsgrense: arbeidsboken må behandles som sensitiv.
    clientSecret = modNettskjema.GetGlobalClientSecret()

    ' Migreringsfallback fra eldre versjoner.
    If clientSecret = "" Then
        clientSecret = modCredentialManager.CredMgr_Read(CRED_TARGET_PREFIX & ClientId)
    End If

    If clientSecret = "" Then
        Err.Raise vbObjectError + 1021, "AcquireAccessToken", "Fant ingen Client Secret i arbeidsboken for Client ID """ & ClientId & """. Åpne Innstillinger og lagre den på nytt."
    End If

    For attempt = 1 To MAX_RETRIES
        If NettskjemaAvbrytHenting Then Err.Raise vbObjectError + 1024, "AcquireAccessToken", "Avbrutt av bruker (Esc)."

        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        http.SetTimeouts 10000, 10000, 15000, 20000
        http.Open "POST", TOKEN_URL, False
        http.SetRequestHeader "Authorization", "Basic " & Base64Encode(ClientId & ":" & clientSecret)
        http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        http.SetRequestHeader "Accept", "application/json"

        sendFailed = False
        On Error Resume Next
        http.Send "grant_type=client_credentials"
        If Err.Number = 18 Then
            NettskjemaAvbrytHenting = True
        End If
        If Err.Number <> 0 Then
            sendFailed = True
            sendErrDescription = Err.Description
            Err.Clear
        End If
        On Error GoTo 0

        If sendFailed Then
            lastErrMsg = "Nettverksfeil: " & sendErrDescription
            modNettskjema.LogEvent "POST token", TOKEN_URL, 0, attempt, "RETRY", lastErrMsg, 0, 0
            If NettskjemaAvbrytHenting Then Err.Raise vbObjectError + 1024, "AcquireAccessToken", "Avbrutt av bruker (Esc)."
            If attempt < MAX_RETRIES Then Application.Wait Now + TimeSerial(0, 0, BackoffSeconds(attempt))

        ElseIf http.Status = 200 Then
            Set json = JsonConverter.ParseJson(DecodeUtf8ResponseBody(http.ResponseBody))
            EnsureTokenCache
            mTokens(ClientId) = CStr(json("access_token"))
            expiresIn = CLng(json("expires_in"))
            ' IKKE TimeSerial(0, 0, sekunder) her - TimeSerial sitt tredje
            ' parameter (sekunder) er Integer-typet internt i VBA (maks
            ' 32767), og expires_in fra en ekte Nettskjema-token er ~86400
            ' (24 timer) - det overskrider Integer-grensen og gir "Runtime
            ' error 6: Overflow" HVER GANG. Legg sekundene til direkte som en
            ' brøkdel av et døgn i stedet (Date-verdier er Double internt,
            ' ingen Integer-begrensning).
            mTokenExpiries(ClientId) = Now + (CDbl(expiresIn - TOKEN_EXPIRY_SAFETY_SECONDS) / 86400#)
            Exit Sub

        ElseIf http.Status >= 500 Then
            lastErrMsg = "Midlertidig serverfeil (HTTP " & http.Status & ")"
            modNettskjema.LogEvent "POST token", TOKEN_URL, http.Status, attempt, "RETRY", lastErrMsg, 0, 0
            If attempt < MAX_RETRIES Then Application.Wait Now + TimeSerial(0, 0, BackoffSeconds(attempt))

        Else
            ' 4xx (f.eks. feil Client ID/Secret) - ingen vits å prøve på
            ' nytt, dette retter ikke seg selv.
            Err.Raise vbObjectError + 1022, "AcquireAccessToken", "OAuth feilet (HTTP " & http.Status & "): " & DecodeUtf8ResponseBody(http.ResponseBody)
        End If
    Next attempt

    Err.Raise vbObjectError + 1023, "AcquireAccessToken", "Ga opp å hente token etter " & MAX_RETRIES & " forsøk mot: " & TOKEN_URL & ". Siste feil: " & lastErrMsg
End Sub

Private Function Base64Encode(ByVal PlainText As String) As String
    Dim bytes() As Byte
    Dim xmlDoc As Object
    Dim node As Object

    bytes = StrConv(PlainText, vbFromUnicode)

    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    Set node = xmlDoc.createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = bytes
    Base64Encode = VBA.Replace(VBA.Replace(node.text, vbCr, ""), vbLf, "")
End Function

Private Function BackoffSeconds(ByVal Attempt As Long) As Long
    Dim s As Long
    s = CLng(2 ^ Attempt)
    If s > MAX_BACKOFF_SECONDS Then s = MAX_BACKOFF_SECONDS
    BackoffSeconds = s
End Function

' WinHttpRequest.ResponseText gjetter selv hvilket tegnsett svaret er i,
' basert på Content-Type-headeren - hvis serveren ikke oppgir "charset=utf-8"
' eksplisitt (Nettskjema-APIet gjor ikke det for x-ndjson/json), kan gjetningen
' bli feil og norske bokstaver (æøå) blir stille korrupte tegn i svar-teksten.
' Denne funksjonen dekoder rå-bytene (ResponseBody) som UTF-8 eksplisitt via
' ADODB.Stream, uavhengig av hva WinHttp selv gjetter.
Private Function DecodeUtf8ResponseBody(ByVal RawBytes As Variant) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary
    stream.Open
    stream.Write RawBytes
    stream.Position = 0
    stream.Type = 2 ' adTypeText
    stream.Charset = "utf-8"
    DecodeUtf8ResponseBody = stream.ReadText
    stream.Close
End Function

' GET mot API-et med automatisk token-fornyelse (401), rate-limit-respekt
' (429/Retry-After) og exponential backoff ved serverfeil (5xx).
Public Function ApiGetWithRetry(ByVal Url As String, ByVal ClientId As String, Optional ByVal AcceptHeader As String = "application/json") As String
    Dim attempt As Long
    Dim http As Object
    Dim waitSeconds As Long
    Dim statusCode As Long
    Dim retryAfterHeader As String
    Dim usedTokenRefresh As Boolean
    Dim sendFailed As Boolean
    Dim sendErrDescription As String

    usedTokenRefresh = False

    For attempt = 1 To MAX_RETRIES
        If NettskjemaAvbrytHenting Then Err.Raise vbObjectError + 1045, "ApiGetWithRetry", "Avbrutt av bruker (Esc)."

        EnsureAccessToken ClientId

        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        ' Receive holdt hoy (120s) - et skjema med mange submissions kan
        ' legitimt ta en stund å generere/overføre server-side. Ikke
        ' forveksle en treg (men fungerende) henting med en genuin
        ' nettverksfeil.
        http.SetTimeouts 10000, 10000, 30000, 120000
        http.Open "GET", Url, False
        http.SetRequestHeader "Authorization", "Bearer " & mTokens(ClientId)
        http.SetRequestHeader "Accept", AcceptHeader
        http.SetRequestHeader "User-Agent", USER_AGENT

        sendFailed = False
        On Error Resume Next
        http.Send
        If Err.Number = 18 Then
            NettskjemaAvbrytHenting = True
        End If
        If Err.Number <> 0 Then
            sendFailed = True
            sendErrDescription = Err.Description
            Err.Clear
        End If
        On Error GoTo 0

        If sendFailed Then
            modNettskjema.LogEvent "GET", Url, 0, attempt, "RETRY", "Nettverksfeil: " & sendErrDescription, 0, 0
            If NettskjemaAvbrytHenting Then Err.Raise vbObjectError + 1045, "ApiGetWithRetry", "Avbrutt av bruker (Esc)."
            If attempt < MAX_RETRIES Then
                waitSeconds = BackoffSeconds(attempt)
                Application.Wait Now + TimeSerial(0, 0, waitSeconds)
            End If
        Else
            statusCode = http.Status

            Select Case statusCode
            Case 200
                ApiGetWithRetry = DecodeUtf8ResponseBody(http.ResponseBody)
                Exit Function

            Case 401
                If usedTokenRefresh Then
                    Err.Raise vbObjectError + 1040, "ApiGetWithRetry", "HTTP 401 selv etter fornyet token. Sjekk Client ID/Secret og at API-klienten har tilgang til skjemaet."
                End If
                ForgetAccessToken ClientId
                usedTokenRefresh = True

            Case 429
                retryAfterHeader = ""
                On Error Resume Next
                retryAfterHeader = http.GetResponseHeader("Retry-After")
                On Error GoTo 0
                If IsNumeric(retryAfterHeader) Then
                    waitSeconds = CLng(retryAfterHeader)
                Else
                    waitSeconds = BackoffSeconds(attempt)
                End If
                If waitSeconds > MAX_BACKOFF_SECONDS Then waitSeconds = MAX_BACKOFF_SECONDS
                modNettskjema.LogEvent "GET", Url, statusCode, attempt, "RETRY", "Rate limited (Retry-After: " & retryAfterHeader & ")", 0, 0
                If attempt < MAX_RETRIES Then Application.Wait Now + TimeSerial(0, 0, waitSeconds)

            Case 500 To 599
                waitSeconds = BackoffSeconds(attempt)
                modNettskjema.LogEvent "GET", Url, statusCode, attempt, "RETRY", "Midlertidig serverfeil", 0, 0
                If attempt < MAX_RETRIES Then Application.Wait Now + TimeSerial(0, 0, waitSeconds)

            Case Else
                Err.Raise vbObjectError + 1041, "ApiGetWithRetry", "HTTP " & statusCode & ": " & DecodeUtf8ResponseBody(http.ResponseBody)
            End Select
        End If
    Next attempt

    Err.Raise vbObjectError + 1042, "ApiGetWithRetry", "Ga opp etter " & MAX_RETRIES & " forsøk mot: " & Url
End Function

' Enkelt GET-kall med KUN ETT forsøk (ingen retry/backoff som ApiGetWithRetry)
' - brukes for "best effort"-kall som ALDRI skal forsinke hovedhentingen
' merkbart hvis de feiler (f.eks. /elements for å auto-utlede
' spørsmålstekst/alternativ-tekst). /elements har vist seg å svare 401
' "Bearer token malformed" periodevis (trolig en forbigående gateway-glipp -
' se feedback-vba-com-test-gotchas i minnet), og et slikt kall skal da bare
' feile raskt og stille, ikke bruke MAX_RETRIES forsøk med voksende
' ventetid på noe som uansett bare er en ekstra finpuss.
Public Function ApiGetSingleAttempt(ByVal Url As String, ByVal ClientId As String, Optional ByVal AcceptHeader As String = "application/json") As String
    EnsureAccessToken ClientId

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    ' Korte tidsavbrudd bevisst (ikke de samme som ApiGetWithRetry sine) -
    ' dette er et "best effort"-kall, og skal ALDRI kunne legge titalls
    ' ekstra sekunder til HVER ENESTE henting hvis dette endepunktet henger.
    ' Total verste-fall-ventetid her er ~10s, ikke ~40s.
    http.SetTimeouts 2000, 2000, 3000, 5000
    http.Open "GET", Url, False
    http.SetRequestHeader "Authorization", "Bearer " & mTokens(ClientId)
    http.SetRequestHeader "Accept", AcceptHeader
    http.SetRequestHeader "User-Agent", USER_AGENT

    ' http.Send kan i seg selv kaste en RÅ WinHttp-COM-feil (f.eks. et
    ' tidsavbrudd) med Excel/Windows sin egen standardtekst (typisk
    ' "Operasjonen ble tidsavbrutt") - fanges her eksplisitt slik at
    ' feilteksten blir konsistent med resten av funksjonen, i stedet for å
    ' stole på at kallerens On Error Resume Next fanger den riktig.
    On Error Resume Next
    http.Send
    If Err.Number <> 0 Then
        Dim sendErrMsg As String
        sendErrMsg = Err.Description
        Err.Clear
        On Error GoTo 0
        Err.Raise vbObjectError + 1044, "ApiGetSingleAttempt", "Nettverksfeil mot " & Url & ": " & sendErrMsg
    End If
    On Error GoTo 0

    If http.Status <> 200 Then
        Err.Raise vbObjectError + 1043, "ApiGetSingleAttempt", "HTTP " & http.Status & " mot: " & Url
    End If

    ApiGetSingleAttempt = DecodeUtf8ResponseBody(http.ResponseBody)
End Function
'@

# ---- modJsonFlatten: parser NDJSON-svaret fra /answers og PIVOTERER det til
# bredt format (en rad per respondent) ----

$flattenCode = @'
Attribute VB_Name = "modJsonFlatten"
Option Explicit

' Parser NDJSON-svaret fra /form/{formId}/answers - som rått er LANGT format
' (en rad per ENKELTSVAR: submissionId, elementId, textAnswer, ...) - og
' PIVOTERER det til BREDT format (en rad per RESPONDENT/submission, en
' kolonne per elementId) før det skrives til arket. Se
' feedback-vba-com-test-gotchas i minnet for bakgrunnen på hvorfor rådataen
' er i langt format.
'
' Kolonnenavn er stabilt "Spm_<elementId>" INTERNT (brukt til å matche
' kolonner på tvers av hentinger og i kolonneinnstillingene), men det som
' faktisk SKRIVES som overskrift er alias hvis Håkon har satt et via
' "Kolonner..."-vinduet, ellers råkey. /form/{formId}/elements svarte 401 for
' denne API-klienten tidlig i prosjektet (trolig en forbigående gateway-glipp
' - se feedback-vba-com-test-gotchas i minnet), men fungerer nå: hvis den
' svarer riktig, auto-utledes alias fra ekte spørsmålstekst (kun når feltet
' er tomt fra før - overskriver aldri en manuell tilpasning), og
' avkrysning/radio-svar vises med ekte alternativ-tekst i stedet for rå
' alternativ-ID-er. Se ElementsInfo-parameteren under.
Public Sub ParseAndWriteSubmissions(ByVal JsonText As String, ByVal FormId As String, ByVal TargetSheetName As String, ByVal TargetTableName As String, Optional ByVal ElementsInfo As Object = Nothing, Optional ByRef NewRowCount As Long)
    Dim submissionRows As Object      ' submissionId (String) -> radens Dictionary
    Dim submissionOrder As Collection ' bevarer rekkefolgen submissions forste gang ble sett
    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    Dim answer As Object

    Set submissionRows = CreateObject("Scripting.Dictionary")
    Set submissionOrder = New Collection

    ' Metadata-kolonnene registreres også i kolonneinnstillingene, synlige som
    ' standard (samme som spørsmålskolonnene) - kan skjules via
    ' "Kolonner..."-dialogen ved behov.
    modNettskjema.EnsureColumnRegistered FormId, "submissionId"
    modNettskjema.EnsureColumnRegistered FormId, "createdDate"
    modNettskjema.EnsureColumnRegistered FormId, "modifiedDate"

    Dim colSubmissionId As String, colCreatedDate As String, colModifiedDate As String
    colSubmissionId = modNettskjema.ResolveColumnDisplayName(FormId, "submissionId")
    colCreatedDate = modNettskjema.ResolveColumnDisplayName(FormId, "createdDate")
    colModifiedDate = modNettskjema.ResolveColumnDisplayName(FormId, "modifiedDate")

    ' Forhåndsregistrerer en kolonne for HVERT ekte spørsmål i selve skjema-
    ' strukturen (ikke bare spørsmål som allerede har fått svar) - løser at
    ' forgreinede/betingede spørsmål er usynlige helt til noen faktisk
    ' svarer. Helt "best effort": ElementsInfo er Nothing hvis /elements
    ' feilet, og da hoppes dette bare over (samme oppførsel som før).
    If Not ElementsInfo Is Nothing Then
        modNettskjema.PreRegisterFormColumns FormId, ElementsInfo
    End If

    lines = VBA.Split(VBA.Replace(JsonText, vbCrLf, vbLf), vbLf)

    For i = LBound(lines) To UBound(lines)
        lineText = VBA.Trim$(lines(i))
        If lineText <> "" Then
            Set answer = JsonConverter.ParseJson(lineText)

            Dim subId As String
            subId = CStr(answer("submissionId"))

            Dim row As Object
            If Not submissionRows.Exists(subId) Then
                Set row = CreateObject("Scripting.Dictionary")
                Set submissionRows(subId) = row
                submissionOrder.Add subId
                row(colSubmissionId) = answer("submissionId")
                If answer.Exists("createdDate") Then row(colCreatedDate) = answer("createdDate")
                If answer.Exists("modifiedDate") Then row(colModifiedDate) = answer("modifiedDate")
            Else
                Set row = submissionRows(subId)
            End If

            Dim elementIdStr As String
            elementIdStr = CStr(answer("elementId"))

            Dim rawKey As String
            rawKey = "Spm_" & elementIdStr
            modNettskjema.EnsureColumnRegistered FormId, rawKey

            ' Auto-utled et lesbart kolonnenavn fra ekte spørsmålstekst hvis
            ' ElementsInfo er tilgjengelig (fra /form/{formId}/elements) OG
            ' kolonnen ikke allerede har et alias fra før - overskriver ALDRI
            ' et alias Håkon selv har satt. Se TryFetchFormElements i
            ' modNettskjema for hvorfor dette er "best effort" og ikke kritisk.
            Dim answerOptionsMap As Object
            Set answerOptionsMap = Nothing
            If Not ElementsInfo Is Nothing Then
                If ElementsInfo.Exists(elementIdStr) Then
                    Dim elInfo As Object
                    Set elInfo = ElementsInfo(elementIdStr)
                    If elInfo.Exists("text") Then
                        modNettskjema.SetColumnAliasIfEmpty FormId, rawKey, CStr(elInfo("text"))
                    End If
                    If elInfo.Exists("answerOptions") Then Set answerOptionsMap = elInfo("answerOptions")
                End If
            End If

            Dim colKey As String
            colKey = modNettskjema.ResolveColumnDisplayName(FormId, rawKey)

            Dim valueText As String
            valueText = AnswerValueText(answer, answerOptionsMap)

            If row.Exists(colKey) Then
                ' Flere svar for samme elementId på samme submission (f.eks.
                ' checkbox med flere valgte alternativer på separate rader) -
                ' slå sammen i stedet for å overskrive.
                If CStr(row(colKey)) <> "" And valueText <> "" Then
                    row(colKey) = CStr(row(colKey)) & "; " & valueText
                ElseIf valueText <> "" Then
                    row(colKey) = valueText
                End If
            Else
                row(colKey) = valueText
            End If
        End If
    Next i

    Dim rowDicts As Collection
    Set rowDicts = New Collection
    Dim sid As Variant
    For Each sid In submissionOrder
        rowDicts.Add submissionRows(sid)
    Next sid

    ' Kolonnelisten hentes fra det REGISTRERTE settet for skjemaet (ikke bare
    ' det som faktisk er svart på i DENNE hentingen) - dette holder
    ' kolonnerekkefølgen stabil på tvers av hentinger (kolonner fjernes aldri,
    ' bare står tomme hvis et spørsmål forsvinner), slik at eksterne formler
    ' andre steder i arbeidsboken som peker på en bestemt celle i tabellen
    ' ikke ødelegges når noen henter data på nytt. Se WriteRowsToSheet.
    Dim finalColumns As Collection
    Set finalColumns = modNettskjema.GetRegisteredColumnDisplayNames(FormId)

    modNettskjema.WriteRowsToSheet rowDicts, finalColumns, TargetSheetName, TargetTableName, NewRowCount
    modNettskjema.ApplyColumnVisibility FormId, TargetSheetName, TargetTableName
End Sub

' textAnswer dekker vanlige tekst-/tall-/dato-svar. answerOptionIds dekker
' avkrysning/radio (flere valgte alternativ-ID-er slås sammen med "; ") - vises
' som ekte alternativ-tekst (f.eks. "9.trinn") hvis AnswerOptionsMap har en
' oppføring for ID-en, ellers faller det tilbake til selve ID-tallet.
Private Function AnswerValueText(ByVal Answer As Object, Optional ByVal AnswerOptionsMap As Object = Nothing) As String
    Dim valueText As String
    valueText = ""

    If Answer.Exists("textAnswer") Then
        If Not IsNull(Answer("textAnswer")) Then valueText = CStr(Answer("textAnswer"))
    End If

    If valueText = "" And Answer.Exists("answerOptionIds") Then
        Dim opts As Object
        On Error Resume Next
        Set opts = Answer("answerOptionIds")
        On Error GoTo 0
        If Not opts Is Nothing Then
            If VBA.TypeName(opts) = "Collection" Then
                Dim k As Long
                For k = 1 To opts.Count
                    Dim optIdStr As String
                    optIdStr = CStr(opts(k))
                    Dim optDisplay As String
                    optDisplay = optIdStr
                    If Not AnswerOptionsMap Is Nothing Then
                        If AnswerOptionsMap.Exists(optIdStr) Then optDisplay = AnswerOptionsMap(optIdStr)
                    End If
                    If valueText <> "" Then valueText = valueText & "; "
                    valueText = valueText & optDisplay
                Next k
            End If
        End If
    End If

    AnswerValueText = valueText
End Function

' Parser /form/{formId}/elements-responsen til et oppslag: elementId (String
' som nøkkel) -> Dictionary med "text" (ekte spørsmålstekst, om satt og ikke
' tom) og "answerOptions" (Dictionary: answerOptionId (String) -> ekte
' alternativ-tekst). Brukes til å auto-utlede lesbare kolonnenavn og
' svar-tekster i stedet for rå element-/alternativ-ID-er - se
' TryFetchFormElements i modNettskjema, som kaller denne og fanger opp
' eventuelle feil (dette er et "best effort"-kall, ikke kritisk).
Public Function ParseElementsJson(ByVal JsonText As String) As Object
    Dim elements As Object
    Set elements = JsonConverter.ParseJson(JsonText)

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim el As Object
    For Each el In elements
        Dim elementIdStr As String
        elementIdStr = CStr(el("elementId"))

        Dim info As Object
        Set info = CreateObject("Scripting.Dictionary")

        If el.Exists("text") Then
            If Not IsNull(el("text")) Then
                Dim txt As String
                txt = VBA.Trim$(CStr(el("text")))
                If txt <> "" Then info("text") = txt
            End If
        End If

        ' "questionId" er ikke-Null KUN for elementer det faktisk er mulig å
        ' svare på (QUESTION, NUMBER, SELECT, NAME, EMAIL osv.) - rene
        ' HEADING/TEXT-elementer (overskrifter/informasjonstekst) har alltid
        ' Null her. Brukes til å forhåndsregistrere kolonner KUN for ekte
        ' spørsmål - se PreRegisterFormColumns i modNettskjema.
        Dim isQuestion As Boolean
        isQuestion = False
        If el.Exists("questionId") Then
            If Not IsNull(el("questionId")) Then isQuestion = True
        End If
        info("isQuestion") = isQuestion

        Dim optionsMap As Object
        Set optionsMap = CreateObject("Scripting.Dictionary")
        If el.Exists("answerOptions") Then
            Dim opts As Object
            On Error Resume Next
            Set opts = el("answerOptions")
            On Error GoTo 0
            If Not opts Is Nothing Then
                If VBA.TypeName(opts) = "Collection" Then
                    Dim k As Long
                    For k = 1 To opts.Count
                        Dim opt As Object
                        Set opt = opts(k)
                        If opt.Exists("answerOptionId") And opt.Exists("text") Then
                            If Not IsNull(opt("text")) Then
                                optionsMap(CStr(opt("answerOptionId"))) = CStr(opt("text"))
                            End If
                        End If
                    Next k
                End If
            End If
        End If
        Set info("answerOptions") = optionsMap

        Set result(elementIdStr) = info
    Next el

    Set ParseElementsJson = result
End Function

'@

# ---- modNettskjema: orkestrering, ark, knapper, logging, versjon ----

$mainCode = @'
Attribute VB_Name = "modNettskjema"
Option Explicit

Public Const NETTSKJEMA_HENTER_VERSION As String = "3.25.0"

' 3.22.1 - Fjernet raden under listen (antall valgt, hjelpetekst, Velg alle/
'          Fjern alle) fra hovedmenyen; listen er høyere i stedet.
'
' 3.22.0 - Nytt grensesnitt:
'          - Hovedmenyen: fargede knapper i logiske grupper, avkrysningsliste
'            (samme markeringsmÃ¸nster som Mail-utsender, inkl. Shift + Velg
'            alle/Fjern alle) og EN knapp "Oppdater markerte nettskjema" i
'            stedet for "Hent svar pÃ¥ nytt" + "Oppdater alle". Ingen Lukk-knapp.
'          - Ny Innstillinger-meny (frmNettskjemaInnstillinger): egen seksjon
'            for Nettskjema API (Client ID, Client Secret, faste adresser,
'            Test tilkobling) + kolonneinnstillinger for alle tilkoblede skjema.
'          - Kolonneinnstillinger Ã¥pnes fra Innstillinger, og lukkes tilbake
'            dit (fÃ¸r lukket den hovedmenyen samtidig).
'          - Endret Client ID synkroniseres til alle tilkoblinger
'            (SaveGlobalApiSettings).
'          - RefreshAllConnections erstattet av RefreshConnectionRows.

' 3.21.0 - Renskriving/effektivisering av UPSERT-koden fra 3.20.x, ingen
'          endring i selve arkitekturen (fortsatt UPSERT, tabellen bygges
'          fortsatt ALDRI om etter forste gang):
'          - WriteRowsToSheet skriver nå en HEL rad i ETT Range-kall i stedet
'            for ett COM-kall per celle (nCols kall) - samme mengde skriving,
'            langt færre COM-kall ved mange kolonner/rader.
'          - VerifyAndRepairTable logger ikke lenger en "OK"-rad til Log-arket
'            ved HVER henting (kun ved faktisk reparasjon) - unngår at Log
'            fylles med tusenvis av identiske rader over tid.
'          - Nytt: engangstilbud (OfferClientSecretMigrationIfNeeded) om å
'            migrere en Client Secret som kun ligger i Windows Credential
'            Manager (fra for 3.20.0) inn i arbeidsboken, slik at portabilitet
'            til andre autoriserte brukere faktisk fungerer uten manuelt
'            oppsett pa hver PC - se notatets pkt. 14.

Public Const SHEET_LOG As String = "Log"
Public Const REGISTRY_SHEET As String = "NettskjemaKoblinger_Hidden"

' Registerkolonner: 1=FormId 2=SheetName 3=TableName 4=ClientId 5=SistOppdatert 6=AntallSvar

Public Const COLUMN_SETTINGS_SHEET As String = "NettskjemaKolonner_Hidden"

' Kolonneinnstillings-kolonner: 1=FormId 2=RawKey (stabil, "Spm_<elementId>")
' 3=Synlig (TRUE/FALSE) 4=Alias (valgfritt visningsnavn)

Private mLastErrorMessage As String
Private mLastNewRowCount As Long

' ---- Inngangspunkt (bundet til knappen) ----

Public Sub AapneNettskjemaHenter()
    OfferClientSecretMigrationIfNeeded
    frmNettskjemaKoblinger.Show
End Sub

' Engangstilbud om å migrere en Client Secret som kun finnes lokalt (Windows
' Credential Manager, fra før 3.20.0) inn i selve arbeidsboken. Uten dette
' kan en kollega som åpner en delt arbeidsbok oppleve at hentingen stille
' feiler, fordi Client Secret aldri ble bedt om på deres PC (se
' AcquireAccessToken sin migreringsfallback i modAuthentication) - se notatet
' "Notat_Nettskjema_henter_oppdatert_3.20.3_v2.txt", pkt. 12-14.
' Bevisst KUN kalt fra AapneNettskjemaHenter (interaktiv åpning av
' tilkoblingsvinduet) - ALDRI fra RefreshConnectionRows/RefreshConnectionRow,
' siden en MsgBox midt i en automatisert henting ville låst en bakgrunnskjøring.
Private Sub OfferClientSecretMigrationIfNeeded()
    Dim clientId As String
    clientId = GetGlobalClientId()
    If clientId = "" Then Exit Sub
    If GetGlobalClientSecret() <> "" Then Exit Sub ' allerede migrert/lagret

    Dim localSecret As String
    localSecret = modCredentialManager.CredMgr_Read(modAuthentication.CRED_TARGET_PREFIX & clientId)
    If localSecret = "" Then Exit Sub ' ingenting å migrere

    If MsgBox("Det finnes en Client Secret lagret lokalt på denne PC-en (Windows Credential Manager), " & _
        "men den er ikke lagret i selve arbeidsboken." & vbCrLf & vbCrLf & _
        "Vil du lagre den i arbeidsboken, slik at Nettskjema Henter kan brukes av andre autoriserte " & _
        "brukere av denne filen uten at de må konfigurere noe selv?" & vbCrLf & vbCrLf & _
        "VIKTIG: Arbeidsboken må da behandles som sensitiv - alle med tilgang til filen kan i " & _
        "prinsippet hente ut secret-en.", vbQuestion + vbYesNo, "Nettskjema Henter - migrer Client Secret") = vbYes Then
        SetGlobalClientSecret localSecret
        MsgBox "Client Secret er nå lagret i arbeidsboken." & vbCrLf & vbCrLf & _
            "Husk å lagre (Ctrl+S) og lukke filen, og distribuer denne lagrede versjonen til kollegaer.", _
            vbInformation, "Nettskjema Henter"
    End If
End Sub

' ---- Tilkoblinger ----

' ---- Generelle innstillinger (ETT felles API-klient-oppsett for alle
' tilkoblinger - Client ID og Client Secret lagres i arbeidsbokens VeryHidden
' registerark slik at samme fil kan brukes av flere autoriserte brukere) ----

Public Sub AapneGenerelleInnstillinger()
    frmNettskjemaInnstillinger.Show
End Sub

' Lagrer det felles API-oppsettet (Client ID + Client Secret) i arbeidsboken.
' Alle tilkoblinger deler ETT oppsett (Ã©n global Client Secret), sÃ¥ Client ID
' i hver registerrad (kolonne 4, brukt av RefreshConnectionRow) holdes i synk
' - ellers ville en endret Client ID gitt henting med gammel ID + ny secret.
Public Sub SaveGlobalApiSettings(ByVal ClientId As String, ByVal ClientSecret As String)
    Dim oldClientId As String
    Dim reg As Worksheet
    Dim lastRow As Long, r As Long

    oldClientId = GetGlobalClientId()

    SetGlobalClientId ClientId
    SetGlobalClientSecret ClientSecret

    Set reg = GetRegistrySheet()
    lastRow = reg.Cells(reg.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        reg.Cells(r, 4).Value = ClientId
    Next r

    ' Tving nytt token neste gang (gammelt token kan tilhÃ¸re gammel ID/secret).
    If oldClientId <> "" Then modAuthentication.ForgetAccessToken oldClientId
    modAuthentication.ForgetAccessToken ClientId
End Sub
Public Function GetGlobalClientId() As String
    GetGlobalClientId = Trim$(CStr(GetRegistrySheet().Range("I1").Value))
End Function

Private Sub SetGlobalClientId(ByVal ClientId As String)
    Dim reg As Worksheet
    Set reg = GetRegistrySheet()
    reg.Range("H1").Value = "GlobalClientId"
    reg.Range("I1").Value = ClientId
End Sub

Public Function GetGlobalClientSecret() As String
    GetGlobalClientSecret = CStr(GetRegistrySheet().Range("I2").Value)
End Function

Private Sub SetGlobalClientSecret(ByVal ClientSecret As String)
    Dim reg As Worksheet
    Set reg = GetRegistrySheet()
    reg.Range("H2").Value = "GlobalClientSecret"
    reg.Range("I2").Value = ClientSecret
End Sub

' ---- Tilkoblinger ----

Public Sub LeggTilTilkobling()
    Dim reg As Worksheet
    Dim formId As String
    Dim clientId As String
    Dim sheetName As String
    Dim existingRow As Long
    Dim existingSheetName As String

    Set reg = GetRegistrySheet()

    clientId = GetGlobalClientId()
    If clientId = "" Then
        If MsgBox("Client ID/Client Secret er ikke satt opp ennå." & vbCrLf & vbCrLf & _
            "Vil du åpne generelle innstillinger nå?", vbQuestion + vbYesNo, "Mangler oppsett") = vbYes Then
            AapneGenerelleInnstillinger
            clientId = GetGlobalClientId()
        End If
        If clientId = "" Then Exit Sub
    End If

    If GetGlobalClientSecret() = "" And _
       modCredentialManager.CredMgr_Read(modAuthentication.CRED_TARGET_PREFIX & clientId) = "" Then
        MsgBox "Client Secret mangler for Client ID """ & clientId & """." & vbCrLf & _
            "Åpne ""Innstillinger..."" for å lagre den i arbeidsboken først.", vbExclamation, "Mangler Client Secret"
        Exit Sub
    End If

    formId = Trim$(InputBox("Form ID for Nettskjema-skjemaet (tallet i nettskjema-URL-en):", "Ny tilkobling"))
    If formId = "" Then Exit Sub
    If Not IsNumeric(formId) Then
        MsgBox "Form ID må være et tall.", vbExclamation
        Exit Sub
    End If

    existingRow = FindConnectionByFormId(formId)
    If existingRow > 0 Then
        existingSheetName = CStr(reg.Cells(existingRow, 2).Value)
        If SheetExists(existingSheetName) Then
            If MsgBox("Dette skjemaet (Form ID " & formId & ") er allerede koblet til arket """ & existingSheetName & """." & vbCrLf & vbCrLf & _
                "Vil du gå til det arket?", vbQuestion + vbYesNo, "Allerede koblet") = vbYes Then
                ThisWorkbook.Worksheets(existingSheetName).Activate
            End If
            Exit Sub
        Else
            If MsgBox("Dette skjemaet var tidligere koblet til arket """ & existingSheetName & """, som ikke lenger finnes i arbeidsboken." & vbCrLf & vbCrLf & _
                "Vil du fjerne den gamle registreringen og lage en ny tilkobling?", vbQuestion + vbYesNo, "Foreldet tilkobling") = vbYes Then
                reg.Rows(existingRow).Delete
            Else
                Exit Sub
            End If
        End If
    End If

    Do
        sheetName = Trim$(InputBox("Navn på arket for dette skjemaet (brukes også til å lage tabellnavnet ""<navn>_tbl""):", "Ny tilkobling"))
        If sheetName = "" Then Exit Sub
        If Not IsValidSheetName(sheetName) Then
            MsgBox "Ugyldig arknavn. Unngå tegnene : \ / ? * [ ] og maks 31 tegn.", vbExclamation
            sheetName = ""
        ElseIf SheetExists(sheetName) Then
            MsgBox "Et ark med dette navnet finnes allerede. Velg et annet navn.", vbExclamation
            sheetName = ""
        End If
    Loop While sheetName = ""

    CreateConnectionSheet formId, clientId, sheetName
End Sub

Private Sub CreateConnectionSheet(ByVal FormId As String, ByVal ClientId As String, ByVal SheetName As String)
    Dim ws As Worksheet
    Dim reg As Worksheet
    Dim tableName As String
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = SheetName
    ws.Range("A1").Value = "Henter svar fra Nettskjema (Form ID " & FormId & ") ..."

    tableName = SanitizeTableName(SheetName) & "_tbl"

    Set reg = GetRegistrySheet()
    r = reg.Cells(reg.Rows.Count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    reg.Cells(r, 1).Value = FormId
    reg.Cells(r, 2).Value = SheetName
    reg.Cells(r, 3).Value = tableName
    reg.Cells(r, 4).Value = ClientId
    reg.Cells(r, 5).Value = ""
    reg.Cells(r, 6).Value = 0

    ws.Activate
    If Not RefreshConnectionRow(r) Then
        MsgBox "Arket og tabellen er opprettet, men første henting feilet:" & vbCrLf & LastErrorMessage() & _
            vbCrLf & vbCrLf & "Bruk ""Hent svar på nytt"" når feilen er rettet.", vbExclamation, "Ny tilkobling"
    End If
End Sub

' Henter (på nytt) alle submissions for en registrert tilkobling - FULL
' REFRESH av tabellen på dennes eget ark. RegistryRow er radnummeret i det
' skjulte registeret (ikke synlig for brukeren).
' Henter (på nytt) EN tilkobling. Viser IKKE MsgBox selv - returnerer True/
' False slik at bade enkelt-oppdatering (som viser sin egen melding) og
' "Oppdater alle" (som viser EN samlet oppsummering til slutt) kan style
' tilbakemeldingen selv. LastErrorMessage() gir feilteksten ved False.
' LastNewRowCount() gir antall NYE submissions (ikke totalt) ved True.
Public Function RefreshConnectionRow(ByVal RegistryRow As Long) As Boolean
    Dim reg As Worksheet
    Dim formId As String, sheetName As String, tableName As String, clientId As String
    Dim jsonText As String
    Dim startTime As Date
    Dim rowCount As Long
    Dim newCount As Long

    Set reg = GetRegistrySheet()
    formId = CStr(reg.Cells(RegistryRow, 1).Value)
    sheetName = CStr(reg.Cells(RegistryRow, 2).Value)
    tableName = CStr(reg.Cells(RegistryRow, 3).Value)
    clientId = CStr(reg.Cells(RegistryRow, 4).Value)

    ' Alt av Esc-status fra en TIDLIGERE rad i samme runde skal hoppe rett
    ' over denne raden også, i stedet for å starte en ny henting brukeren
    ' allerede har bedt om å stoppe (se modAuthentication.NettskjemaAvbrytHenting).
    If modAuthentication.NettskjemaAvbrytHenting Then
        mLastErrorMessage = "Avbrutt av bruker (Esc)."
        RefreshConnectionRow = False
        Exit Function
    End If

    startTime = Now
    On Error GoTo ErrHandler

    Application.Cursor = xlWait
    Application.StatusBar = "Henter submissions for " & sheetName & " (Form ID " & formId & ") ..."

    ' MERK: "/form/{formId}/submissions" (flertall) finnes IKKE i det ekte v3-
    ' API-et - det var v2-APIets endepunkt, feilaktig videreført fra det
    ' opprinnelige idé-notatet uten å sjekkes mot den faktiske Swagger-
    ' dokumentasjonen. "/form/{formId}/answers" er riktig v3-endepunkt for
    ' å hente alle svar for et skjema (bekreftet: "/submissions" gir
    ' HTTP 405 "Request method 'GET' is not supported").
    ' Dette endepunktet svarer KUN med application/x-ndjson (bekreftet: vanlig
    ' application/json gir HTTP 406 "No acceptable representation") - hver
    ' linje i svaret er sin egen JSON-verdi, IKKE én stor JSON-array.
    jsonText = modAuthentication.ApiGetWithRetry(modAuthentication.API_BASE_URL & "/form/" & formId & "/answers", clientId, "application/x-ndjson")

    Dim elementsInfo As Object
    Set elementsInfo = TryFetchFormElements(formId, clientId)

    modJsonFlatten.ParseAndWriteSubmissions jsonText, formId, sheetName, tableName, elementsInfo, newCount

    rowCount = GetTableRowCount(sheetName, tableName)
    reg.Cells(RegistryRow, 5).Value = Now
    reg.Cells(RegistryRow, 6).Value = rowCount

    LogEvent "GET submissions", sheetName & " (Form " & formId & ")", 200, 1, "SUCCESS", "", rowCount, DateDiff("s", startTime, Now)

    Application.StatusBar = False
    Application.Cursor = xlDefault
    mLastErrorMessage = ""
    mLastNewRowCount = newCount
    RefreshConnectionRow = True
    Exit Function

ErrHandler:
    ' MAA leses FORST, før noe annet kjores: LogEvent kaller GetLogSheet, som
    ' bruker "On Error Resume Next"/"On Error GoTo 0" internt - ethvert
    ' "On Error"-statement NULLSTILLER Err-objektet automatisk (dokumentert
    ' VBA-oppforsel), så Err.Description ville blitt tom hvis den ble lest
    ' ETTER LogEvent-kallet under.
    Dim errMsg As String
    errMsg = Err.Description

    Application.StatusBar = False
    Application.Cursor = xlDefault
    LogEvent "GET submissions", sheetName & " (Form " & formId & ")", 0, 0, "ERROR", errMsg, 0, DateDiff("s", startTime, Now)
    mLastErrorMessage = errMsg
    mLastNewRowCount = 0
    RefreshConnectionRow = False
End Function

Public Function LastErrorMessage() As String
    LastErrorMessage = mLastErrorMessage
End Function

' Antall NYE submissions (fantes ikke i tabellen fra før) fra siste
' RefreshConnectionRow-kall - 0 ved feil. Brukt av RefreshConnectionRows for
' å vise "X nye svar" i oppsummeringen, uten å måtte regne det ut på nytt.
Public Function LastNewRowCount() As Long
    LastNewRowCount = mLastNewRowCount
End Function

' Fjerner en tilkobling helt: sletter arket + tabellen (all data i den) OG
' registerraden. Kalles kun etter at brukeren har bekreftet i en advarsel
' (se cmdFjernTilkobling_Click i frmNettskjemaKoblinger) - ingen egen
' bekreftelse her.
Public Sub RemoveConnection(ByVal RegistryRow As Long)
    Dim reg As Worksheet
    Dim formId As String
    Dim sheetName As String

    Set reg = GetRegistrySheet()
    formId = CStr(reg.Cells(RegistryRow, 1).Value)
    sheetName = CStr(reg.Cells(RegistryRow, 2).Value)

    If SheetExists(sheetName) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(sheetName).Delete
        Application.DisplayAlerts = True
    End If

    reg.Rows(RegistryRow).Delete

    ' Fjerner ALT tilknyttet dette skjemaet, ikke bare arket/tabellen og
    ' registerraden - uten dette ville et skjema som fjernes og legges inn på
    ' nytt (med samme Form ID) gjenbruke gammel kolonnerekkefølge og gamle
    ' aliaser fra det skjulte kolonneinnstillingsarket i stedet for å starte
    ' helt friskt.
    RemoveColumnSettingsForForm formId
End Sub

' Fjerner alle registrerte kolonneinnstillinger (synlig/skjult + alias) for
' et gitt Form ID - se RemoveConnection.
Private Sub RemoveColumnSettingsForForm(ByVal FormId As String)
    Dim sh As Worksheet
    Set sh = GetColumnSettingsSheet()
    Dim lastRow As Long, r As Long
    lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    For r = lastRow To 2 Step -1
        If CStr(sh.Cells(r, 1).Value) = FormId Then
            sh.Rows(r).Delete
        End If
    Next r
End Sub

' ---- Kolonneinnstillinger (synlig/skjult + alias) per tilkobling ----
' Delt på tvers av ALLE tilkoblinger i arbeidsboken (ett skjult ark), radene
' filtreres på FormId. RawKey ("Spm_<elementId>") er den STABILE nøkkelen -
' den endrer seg aldri, selv om Håkon endrer et alias - dette er bevisst,
' slik at synlig/skjult-status og alias overlever selv om aliaset endres.

' Rent lagringsark - vises ALDRI direkte til Håkon (Very Hidden hele tiden).
' Redigeres kun via frmKolonneInnstillinger-dialogen.
Public Function GetColumnSettingsSheet() As Worksheet
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(COLUMN_SETTINGS_SHEET)
    On Error GoTo 0
    If sh Is Nothing Then
        Set sh = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        sh.Name = COLUMN_SETTINGS_SHEET
        sh.Range("A1:D1").Value = Array("FormId", "RawKey", "Synlig", "Alias")
        sh.Rows(1).Font.Bold = True
        sh.Visible = xlSheetVeryHidden
    End If
    Set GetColumnSettingsSheet = sh
End Function

Private Function FindColumnSettingsRow(ByVal FormId As String, ByVal RawKey As String) As Long
    Dim sh As Worksheet
    Dim lastRow As Long, r As Long
    Set sh = GetColumnSettingsSheet()
    lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If CStr(sh.Cells(r, 1).Value) = FormId And CStr(sh.Cells(r, 2).Value) = RawKey Then
            FindColumnSettingsRow = r
            Exit Function
        End If
    Next r
    FindColumnSettingsRow = 0
End Function

' Sikrer at en (FormId, RawKey)-kombinasjon finnes i innstillingsarket -
' kalles for HVER kolonne som oppdages under en henting, slik at den dukker
' opp i "Kolonner..."-vinduet med standardverdier (synlig, uten alias) med
' en gang, også før Håkon har rukket å tilpasse noe.
Public Sub EnsureColumnRegistered(ByVal FormId As String, ByVal RawKey As String, Optional ByVal DefaultVisible As Boolean = True)
    If FindColumnSettingsRow(FormId, RawKey) > 0 Then Exit Sub
    Dim sh As Worksheet
    Set sh = GetColumnSettingsSheet()
    Dim newRow As Long
    newRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row + 1
    If newRow < 2 Then newRow = 2
    sh.Cells(newRow, 1).Value = FormId
    sh.Cells(newRow, 2).Value = RawKey
    sh.Cells(newRow, 3).Value = DefaultVisible
    sh.Cells(newRow, 4).Value = ""
End Sub

' Returnerer alias hvis satt, ellers RawKey uendret.
Public Function ResolveColumnDisplayName(ByVal FormId As String, ByVal RawKey As String) As String
    Dim r As Long
    r = FindColumnSettingsRow(FormId, RawKey)
    If r = 0 Then
        ResolveColumnDisplayName = RawKey
        Exit Function
    End If
    Dim sh As Worksheet
    Set sh = GetColumnSettingsSheet()
    Dim aliasVal As String
    aliasVal = CStr(sh.Cells(r, 4).Value)
    If aliasVal <> "" Then
        ResolveColumnDisplayName = aliasVal
    Else
        ResolveColumnDisplayName = RawKey
    End If
End Function

' Setter alias KUN hvis feltet er tomt fra før - overskriver ALDRI et alias
' Håkon selv har satt manuelt via "Kolonner..."-dialogen. Brukes til å
' auto-utlede et lesbart kolonnenavn fra ekte spørsmålstekst (se
' TryFetchFormElements) uten å ødelegge egne tilpasninger.
Public Sub SetColumnAliasIfEmpty(ByVal FormId As String, ByVal RawKey As String, ByVal ProposedAlias As String)
    If VBA.Trim$(ProposedAlias) = "" Then Exit Sub
    Dim r As Long
    r = FindColumnSettingsRow(FormId, RawKey)
    If r = 0 Then Exit Sub
    Dim sh As Worksheet
    Set sh = GetColumnSettingsSheet()
    If CStr(sh.Cells(r, 4).Value) = "" Then
        sh.Cells(r, 4).Value = ProposedAlias
    End If
End Sub

' Forhåndsregistrerer en kolonne for HVERT ekte spørsmål i selve
' skjema-STRUKTUREN (fra /form/{formId}/elements) - ikke bare spørsmål som
' allerede har fått minst ett svar. Løser at forgreinede/betingede spørsmål
' (som ingen har trigget/svart på ennå) var usynlige helt til noen faktisk
' svarte, siden kolonnelisten tidligere bare ble bygget fra observerte svar.
' Kalles FØR selve svar-linjene parses (i ParseAndWriteSubmissions), slik at
' rekkefølgen for et helt nytt skjema følger skjemaets egen spørsmålsrekke-
' følge - kolonner som allerede er registrert fra før (EnsureColumnRegistered
' er idempotent) berøres ikke, nye spørsmål legges kun til på slutten.
' HEADING/TEXT-elementer (rene overskrifter/infotekst - se "isQuestion" i
' ParseElementsJson) hoppes bevisst over, siden de aldri kan ha svar.
Public Sub PreRegisterFormColumns(ByVal FormId As String, ByVal ElementsInfo As Object)
    If ElementsInfo Is Nothing Then Exit Sub

    Dim elementIdVariant As Variant
    For Each elementIdVariant In ElementsInfo.Keys
        Dim elInfo As Object
        Set elInfo = ElementsInfo(elementIdVariant)

        Dim isQuestion As Boolean
        isQuestion = False
        If elInfo.Exists("isQuestion") Then isQuestion = elInfo("isQuestion")
        If Not isQuestion Then GoTo NextElement

        Dim rawKey As String
        rawKey = "Spm_" & CStr(elementIdVariant)
        EnsureColumnRegistered FormId, rawKey

        If elInfo.Exists("text") Then
            SetColumnAliasIfEmpty FormId, rawKey, CStr(elInfo("text"))
        End If
NextElement:
    Next elementIdVariant
End Sub

' Henter /form/{formId}/elements "best effort" - brukes til å auto-utlede
' ekte spørsmålstekst og svaralternativ-tekst i stedet for rå element-/
' alternativ-ID-er. Returnerer Nothing (ALDRI en feil) hvis kallet feiler av
' noen som helst grunn - dette skal ALDRI hindre selve hentingen av svar, kun
' være en finpuss når det er tilgjengelig. Se ApiGetSingleAttempt i
' modAuthentication for hvorfor dette bruker ett forsøk, ikke full retry.
Private Function TryFetchFormElements(ByVal FormId As String, ByVal ClientId As String) As Object
    On Error Resume Next
    Dim jsonText As String
    jsonText = modAuthentication.ApiGetSingleAttempt(modAuthentication.API_BASE_URL & "/form/" & FormId & "/elements", ClientId, "application/json")
    If Err.Number <> 0 Then
        Dim fetchErrMsg As String
        fetchErrMsg = Err.Description
        Err.Clear
        On Error GoTo 0
        LogEvent "GET elements", FormId, 0, 0, "ADVARSEL", "Klarte ikke hente /elements (spørsmålstekst/alternativ-tekst) - faller tilbake til rå ID-er: " & fetchErrMsg, 0, 0
        Set TryFetchFormElements = Nothing
        Exit Function
    End If
    On Error GoTo 0

    Dim result As Object
    On Error Resume Next
    Set result = modJsonFlatten.ParseElementsJson(jsonText)
    If Err.Number <> 0 Then
        Dim parseErrMsg As String
        parseErrMsg = Err.Description
        Err.Clear
        On Error GoTo 0
        LogEvent "GET elements", FormId, 0, 0, "ADVARSEL", "Klarte ikke tolke /elements-svaret - faller tilbake til rå ID-er: " & parseErrMsg, 0, 0
        Set TryFetchFormElements = Nothing
        Exit Function
    End If
    On Error GoTo 0

    LogEvent "GET elements", FormId, 200, 1, "SUCCESS", "", result.Count, 0

    Set TryFetchFormElements = result
End Function

' Returnerer alle registrerte kolonners visningsnavn (alias om satt, ellers
' RawKey) for et gitt skjema, i REGISTRERINGSREKKEFØLGEN. Kolonner fjernes
' aldri fra denne lista selv om et spørsmål slutter å dukke opp i en senere
' henting - de blir bare stående tomme. Dette holder kolonnerekkefølgen
' STABIL på tvers av hentinger, slik at WriteRowsToSheet kan oppdatere
' tabellen i tidligere brukte celler i stedet for å bygge den helt om, og
' eksterne formler andre steder i arbeidsboken som peker på en bestemt celle
' i tabellen ikke blir til #REF! hver gang noen henter data på nytt.
Public Function GetRegisteredColumnDisplayNames(ByVal FormId As String) As Collection
    Dim sh As Worksheet
    Set sh = GetColumnSettingsSheet()
    Dim lastRow As Long, r As Long
    lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    Dim result As New Collection
    For r = 2 To lastRow
        If CStr(sh.Cells(r, 1).Value) = FormId Then
            Dim rawKeyHere As String, aliasHere As String
            rawKeyHere = CStr(sh.Cells(r, 2).Value)
            aliasHere = CStr(sh.Cells(r, 4).Value)
            result.Add IIf(aliasHere <> "", aliasHere, rawKeyHere)
        End If
    Next r
    Set GetRegisteredColumnDisplayNames = result
End Function

' Skjuler/viser de faktiske kolonnene i tabellen basert på lagrede
' innstillinger. Kalles etter hver henting (i ParseAndWriteSubmissions), slik
' at skjulte kolonner FORBLIR skjult også etter "Hent svar på nytt" -
' Excel ville ellers vist alt på nytt siden tabellen bygges helt om for
' hver henting.
Public Sub ApplyColumnVisibility(ByVal FormId As String, ByVal SheetName As String, ByVal TableName As String)
    On Error Resume Next
    Dim ws As Worksheet, tbl As ListObject, col As ListColumn
    Set ws = ThisWorkbook.Worksheets(SheetName)
    Set tbl = ws.ListObjects(TableName)
    If tbl Is Nothing Then Exit Sub

    Dim sh As Worksheet
    Set sh = GetColumnSettingsSheet()
    Dim lastRow As Long, r As Long
    lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row

    For Each col In tbl.ListColumns
        Dim isVisible As Boolean
        isVisible = True
        For r = 2 To lastRow
            If CStr(sh.Cells(r, 1).Value) = FormId Then
                Dim rawKeyHere As String, aliasHere As String, displayNameHere As String
                rawKeyHere = CStr(sh.Cells(r, 2).Value)
                aliasHere = CStr(sh.Cells(r, 4).Value)
                displayNameHere = IIf(aliasHere <> "", aliasHere, rawKeyHere)
                If displayNameHere = col.Name Then
                    isVisible = (sh.Cells(r, 3).Value <> False)
                    Exit For
                End If
            End If
        Next r
        col.Range.EntireColumn.Hidden = Not isVisible
    Next col
End Sub

' Aapner kolonneinnstillings-DIALOGEN (frmKolonneInnstillinger) for en
' tilkobling - ikke et synlig ark. Dialogen bygger sine egne avkrysningsbokser
' + aliasfelt dynamisk (samme mønster som Kolonnevelger sin frmColumnPicker).
Public Sub AapneKolonneinnstillinger(ByVal FormId As String)
    Dim frm As New frmKolonneInnstillinger
    frm.FormId = FormId
    frm.LoadColumnSettings
    frm.Show
End Sub

' Radnummeret i innstillingsarket for en gitt (FormId, RawKey) - brukt av
' dialogen til å lese/skrive verdier direkte. 0 hvis ikke funnet.
Public Function FindColumnSettingsRowPublic(ByVal FormId As String, ByVal RawKey As String) As Long
    FindColumnSettingsRowPublic = FindColumnSettingsRow(FormId, RawKey)
End Function

' Oppdaterer de GITTE tilkoblingene (registerradnumre) etter tur, og viser EN
' samlet oppsummering til slutt (i stedet for en feilboks per tilkobling midt
' i loopen). Gir Håkon full kontroll over når/hvor ofte ny data hentes -
' ingen automatisk bakgrunnsoppdatering noe sted i verktoyet.
Public Sub RefreshConnectionRows(ByVal RegistryRows As Collection)
    Dim reg As Worksheet
    Dim rowNr As Variant
    Dim successCount As Long, failCount As Long
    Dim totalNewRows As Long
    Dim failedList As String
    Dim msg As String

    If RegistryRows Is Nothing Then Exit Sub
    If RegistryRows.Count = 0 Then Exit Sub

    Set reg = GetRegistrySheet()
    modAuthentication.NettskjemaNullstillAvbrudd

    For Each rowNr In RegistryRows
        If RefreshConnectionRow(CLng(rowNr)) Then
            successCount = successCount + 1
            totalNewRows = totalNewRows + LastNewRowCount()
        Else
            failCount = failCount + 1
            failedList = failedList & reg.Cells(CLng(rowNr), 2).Value & " (" & LastErrorMessage() & ")" & vbCrLf
            ' Bruker trykket Esc - hopp over resten av de valgte skjemaene i
            ' stedet for å kjøre gjennom alle likevel (se g-nettskjema-http-
            ' synkron-blokkering); de som faktisk rakk å bli hentet før det,
            ' står fortsatt i successCount/failedList som normalt.
            If modAuthentication.ErHentingAvbrutt() Then Exit For
        End If
    Next rowNr

    msg = successCount & " av " & RegistryRows.Count & " skjema oppdatert. " & totalNewRows & " nye svar."
    If failCount > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Feilet:" & vbCrLf & failedList
    End If
    MsgBox msg, IIf(failCount > 0, vbExclamation, vbInformation), "Oppdater markerte nettskjema"
End Sub
' Bruker /form/me (IKKE /me) - /me svarer 200 med isAuthenticated:false selv
' UTEN token, så den bekrefter ingenting. /form/me krever faktisk gyldig
' token (401 uten), og gir i tillegg en reell sjekk på at API-klienten har
' tilgang til minst ett skjema.
Public Sub TestTilkobling(ByVal ClientId As String)
    On Error GoTo ErrHandler
    Dim jsonText As String
    jsonText = modAuthentication.ApiGetWithRetry(modAuthentication.API_BASE_URL & "/form/me", ClientId)
    MsgBox "OAuth og Nettskjema API fungerer." & vbCrLf & vbCrLf & _
           "Svar fra /form/me (forkortet):" & vbCrLf & VBA.Left$(jsonText, 500), vbInformation, "Test tilkobling"
    Exit Sub
ErrHandler:
    MsgBox "Tilkobling feilet:" & vbCrLf & Err.Description, vbCritical, "Test tilkobling"
End Sub

' ---- Registerhjelpere (skjult ark - se REGISTRY_SHEET) ----

Public Function GetRegistrySheet() As Worksheet
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(REGISTRY_SHEET)
    On Error GoTo 0
    If sh Is Nothing Then
        Set sh = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        sh.Name = REGISTRY_SHEET
        sh.Range("A1:F1").Value = Array("FormId", "SheetName", "TableName", "ClientId", "SistOppdatert", "AntallSvar")
    End If

    ' Registeret inneholder også Client Secret. VeryHidden hindrer tilfeldig
    ' innsyn, men er ikke kryptering eller en sikkerhetsgrense.
    sh.Visible = xlSheetVeryHidden
    If CStr(sh.Range("H1").Value) = "" Then sh.Range("H1").Value = "GlobalClientId"
    If CStr(sh.Range("H2").Value) = "" Then sh.Range("H2").Value = "GlobalClientSecret"
    Set GetRegistrySheet = sh
End Function

Public Function FindConnectionByFormId(ByVal FormId As String) As Long
    Dim reg As Worksheet
    Dim lastRow As Long, r As Long
    Set reg = GetRegistrySheet()
    lastRow = reg.Cells(reg.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If CStr(reg.Cells(r, 1).Value) = FormId Then
            FindConnectionByFormId = r
            Exit Function
        End If
    Next r
    FindConnectionByFormId = 0
End Function

' Teller faktiske submissions via første kolonne (submissionId).
' Beholdes som CountA slik at radtelling er robust også for eldre arbeidsbøker
' som kan ha blitt opprettet med kapasitetsbuffer av en tidligere versjon.
Private Function GetTableRowCount(ByVal SheetName As String, ByVal TableName As String) As Long
    On Error Resume Next
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets(SheetName).ListObjects(TableName)
    If tbl Is Nothing Then Exit Function
    If tbl.ListRows.Count = 0 Then Exit Function
    GetTableRowCount = Application.WorksheetFunction.CountA(tbl.ListColumns(1).DataBodyRange)
    On Error GoTo 0
End Function

' ---- Skriving til et tilkoblings-ark (kalt fra modJsonFlatten) ----

Private Sub VerifyAndRepairTable(ByVal tbl As ListObject, ByVal TargetSheetName As String)
    Dim repaired As Boolean
    Dim expectedBottomRow As Long
    Dim actualBottomRow As Long

    If tbl Is Nothing Then Exit Sub

    If tbl.ListRows.Count > 0 Then
        If tbl.DataBodyRange Is Nothing Then
            Err.Raise vbObjectError + 1064, "VerifyAndRepairTable", _
                "Tabellen '" & tbl.Name & "' har ListRows, men mangler DataBodyRange."
        End If
        If tbl.DataBodyRange.Rows.Count <> tbl.ListRows.Count Then
            Err.Raise vbObjectError + 1065, "VerifyAndRepairTable", _
                "Tabellstrukturen er inkonsistent i '" & tbl.Name & "'."
        End If
        expectedBottomRow = tbl.HeaderRowRange.Row + tbl.ListRows.Count
        actualBottomRow = tbl.DataBodyRange.Row + tbl.DataBodyRange.Rows.Count - 1
        If actualBottomRow <> expectedBottomRow Then
            Err.Raise vbObjectError + 1066, "VerifyAndRepairTable", _
                "Siste datarad i '" & tbl.Name & "' ligger ikke der ListObject-et forventer."
        End If
    End If

    ' Bevar eksisterende tabellstil urørt. Vi skal bare sikre at Excel sin
    ' innebygde radstriping er slått på. Å tilordne TableStyle på nytt her kan
    ' føre til at Excel flater ut/rekalkulerer den visuelle stripingen.
    If Not tbl.ShowTableStyleRowStripes Then
        tbl.ShowTableStyleRowStripes = True
        repaired = True
    End If

    ' Logger KUN når noe faktisk ble reparert - en stille OK-bekreftelse ved
    ' HVER ENESTE henting ville fylt Log-arket med tusenvis av identiske
    ' rader over tid (hentes flere ganger daglig). Ekte strukturfeil stopper
    ' uansett hele hentingen via Err.Raise over, og havner i Log som ERROR.
    If repaired Then
        modNettskjema.LogEvent "Tabellkontroll", TargetSheetName, 0, 0, "REPARERT", _
            "Stripete rader ble reaktivert for tabellen '" & tbl.Name & "'.", tbl.ListRows.Count, 0
    End If
End Sub

Public Sub WriteRowsToSheet(ByVal RowDicts As Collection, ByVal Columns As Collection, ByVal TargetSheetName As String, ByVal TargetTableName As String, Optional ByRef NewRowCount As Long)
    ' Oppdaterer Nettskjema-tabellen IN PLACE. Tabellen Unlist-es, slettes eller
    ' bygges ALDRI opp på nytt når den først finnes. Dermed beholder
    ' NS_SK_tbl/NS_DS_tbl samme ListObject-identitet, og eksterne strukturerte
    ' referanser fortsetter å peke på selve tabellen.
    '
    ' API-kallet er fortsatt en full henting. submissionId brukes som
    ' primærnøkkel (UPSERT): eksisterende submission oppdateres på sin rad,
    ' ny submission legges til som en verifisert ListRow INNE i tabellen.
    ' Eksisterende rader som ikke finnes i et senere API-svar slettes ikke
    ' automatisk; dette er bevisst for
    ' å beskytte mot midlertidig tomme/ufullstendige API-svar.

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim nRows As Long, nCols As Long
    Dim r As Long, c As Long
    Dim colName As Variant
    Dim rowDict As Object
    Dim v As Variant

    Set ws = ThisWorkbook.Worksheets(TargetSheetName)
    nRows = RowDicts.Count
    nCols = Columns.Count

    If nCols = 0 Then Exit Sub

    On Error Resume Next
    Set tbl = ws.ListObjects(TargetTableName)
    If tbl Is Nothing And ws.ListObjects.Count > 0 Then Set tbl = ws.ListObjects(1)
    On Error GoTo 0

    ' Et tomt API-svar skal aldri rive ned eller tømme en eksisterende tabell.
    If nRows = 0 Then
        If Not tbl Is Nothing Then
            modNettskjema.LogEvent "GET submissions", TargetSheetName, 0, 0, "ADVARSEL", _
                "Hentingen ga 0 rader. Eksisterende tabell beholdes uendret.", 0, 0
        Else
            ws.Range("A1").Value = "Ingen submissions funnet for dette skjemaet."
        End If
        Exit Sub
    End If

    If tbl Is Nothing Then
        ' Kun første gang: bygg tabellen med de radene som faktisk finnes.
        ' Senere oppdateringer skjer alltid mot dette samme ListObject-et.
        Dim arr() As Variant
        ReDim arr(0 To nRows, 0 To nCols - 1)

        c = 0
        For Each colName In Columns
            arr(0, c) = colName
            c = c + 1
        Next colName

        r = 1
        For Each rowDict In RowDicts
            c = 0
            For Each colName In Columns
                If rowDict.Exists(colName) Then
                    v = rowDict(colName)
                    If IsEmpty(v) Then v = ""
                    arr(r, c) = v
                Else
                    arr(r, c) = ""
                End If
                c = c + 1
            Next colName
            r = r + 1
        Next rowDict

        ws.Cells.Clear
        ws.Range(ws.Cells(1, 1), ws.Cells(nRows + 1, nCols)).Value = arr
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(nRows + 1, nCols)), , xlYes)
        tbl.Name = TargetTableName
        ws.Columns.AutoFit
        NewRowCount = nRows
        Exit Sub
    End If

    ' Sørg for at tabellen har alle registrerte kolonner. Nye spørsmål kan
    ' derfor legges til uten Unlist/rebuild. Eksisterende ListObject beholdes.
    Do While tbl.ListColumns.Count < nCols
        tbl.ListColumns.Add
    Loop

    ' Hold overskriftene synkronisert med registrert kolonnerekkefølge/alias.
    ' Dette er normalt en no-op. Ved aliasendring får Excel selv oppdatere
    ' strukturerte referanser fordi det samme ListObject-et fortsatt eksisterer.
    c = 1
    For Each colName In Columns
        If CStr(tbl.HeaderRowRange.Cells(1, c).Value) <> CStr(colName) Then
            tbl.HeaderRowRange.Cells(1, c).Value = CStr(colName)
        End If
        c = c + 1
    Next colName

    ' submissionId er første registrerte metadata-kolonne og fungerer som
    ' stabil primærnøkkel for UPSERT-en.
    Dim idColumnName As String
    idColumnName = CStr(Columns(1))

    Dim existingRows As Object
    Set existingRows = CreateObject("Scripting.Dictionary")
    existingRows.CompareMode = vbTextCompare

    If Not tbl.DataBodyRange Is Nothing Then
        For r = 1 To tbl.ListRows.Count
            Dim existingId As String
            existingId = CStr(tbl.DataBodyRange.Cells(r, 1).Value)
            If existingId <> "" Then
                If Not existingRows.Exists(existingId) Then existingRows.Add existingId, r
            End If
        Next r
    End If

    Dim submissionId As String
    Dim targetRow As Long
    Dim lr As ListRow

    For Each rowDict In RowDicts
        If Not rowDict.Exists(idColumnName) Then
            Err.Raise vbObjectError + 1060, "WriteRowsToSheet", _
                "En submission mangler submissionId-kolonnen '" & idColumnName & "'."
        End If

        submissionId = CStr(rowDict(idColumnName))

        If existingRows.Exists(submissionId) Then
            targetRow = CLng(existingRows(submissionId))
        Else
            ' Legg alltid nye submissions eksplisitt INNE i ListObject-et.
            ' AlwaysInsert:=True gjør at Excel flytter eventuell innholdsraden
            ' under tabellen ned i stedet for å risikere en løs rad utenfor.
            Dim rowsBefore As Long
            rowsBefore = tbl.ListRows.Count

            Set lr = tbl.ListRows.Add(AlwaysInsert:=True)

            ' Ikke stol på at Excel "ser riktig ut": verifiser at ListObject-et
            ' faktisk fikk én ny datarad før vi skriver noen verdier.
            If tbl.ListRows.Count <> rowsBefore + 1 Then
                Err.Raise vbObjectError + 1061, "WriteRowsToSheet", _
                    "Klarte ikke å utvide tabellen '" & TargetTableName & _
                    "' med en ny rad. Oppdateringen er avbrutt for å unngå data utenfor tabellen."
            End If

            targetRow = lr.Index

            If tbl.DataBodyRange Is Nothing Then
                Err.Raise vbObjectError + 1062, "WriteRowsToSheet", _
                    "Tabellen '" & TargetTableName & "' mangler DataBodyRange etter at ny rad ble lagt til."
            End If

            If targetRow < 1 Or targetRow > tbl.ListRows.Count Then
                Err.Raise vbObjectError + 1063, "WriteRowsToSheet", _
                    "Ny rad ble ikke registrert som en gyldig ListRow i tabellen '" & TargetTableName & "'."
            End If

            existingRows.Add submissionId, targetRow
            NewRowCount = NewRowCount + 1
        End If

        ' Full henting betyr at raden kan skrives komplett på nytt. Dette gjør
        ' også at et tidligere svar som senere er tømt, faktisk blir tomt.
        '
        ' Skriver HELE raden i ETT område-kall i stedet for ett COM-kall per
        ' celle (nCols kall) - samme prinsipp som det opprinnelige notatet
        ' advarte mot ("ikke skriv celle for celle"). En Variant-array-verdi
        ' satt til Empty gir en EKTE tom celle ved områdetildeling (samme
        ' effekt som ClearContents), så oppførselen er uendret - bare
        ' antall COM-kall er redusert.
        Dim rowArr() As Variant
        ReDim rowArr(1 To 1, 1 To nCols)
        For c = 1 To nCols
            colName = Columns(c)
            If rowDict.Exists(CStr(colName)) Then
                v = rowDict(CStr(colName))
                If IsEmpty(v) Then v = ""
                rowArr(1, c) = v
            Else
                rowArr(1, c) = Empty
            End If
        Next c
        ws.Range(tbl.DataBodyRange.Cells(targetRow, 1), tbl.DataBodyRange.Cells(targetRow, nCols)).Value = rowArr
    Next rowDict

    VerifyAndRepairTable tbl, TargetSheetName
    ws.Columns.AutoFit
End Sub

' ---- Logging (selvhelbredende - oppretter Log-arket ved behov) ----

Public Sub LogEvent(ByVal Operation As String, ByVal FormIdOrUrl As String, ByVal HttpStatus As Long, ByVal Attempt As Long, ByVal Result As String, ByVal Message As String, ByVal RowsAffected As Long, ByVal DurationSeconds As Long)
    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = GetLogSheet()

    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 1).NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ws.Cells(nextRow, 2).Value = Operation
    ws.Cells(nextRow, 3).Value = FormIdOrUrl
    ws.Cells(nextRow, 4).Value = HttpStatus
    ws.Cells(nextRow, 5).Value = Attempt
    ws.Cells(nextRow, 6).Value = Result
    ws.Cells(nextRow, 7).Value = Message
    ws.Cells(nextRow, 8).Value = RowsAffected
    ws.Cells(nextRow, 9).Value = DurationSeconds

    ' Uten dette viser Excel "#####" ("overflow") for Tidspunkt-kolonnen -
    ' standard kolonnebredde er for smal til å vise en full dato+klokkeslett.
    ' Kjores etter HVER rad (ikke bare ved arkoppretting) siden AutoFit basert
    ' kun på header-teksten ved fersk oppretting ikke er bred nok for selve
    ' datoverdiene som kommer senere.
    ws.Columns.AutoFit
End Sub

Private Function GetLogSheet() As Worksheet
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(SHEET_LOG)
    On Error GoTo 0
    If sh Is Nothing Then
        Set sh = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        sh.Name = SHEET_LOG
        sh.Range("A1").Resize(1, 9).Value = Array("Tidspunkt", "Operasjon", "Form/URL", "HTTP status", "Forsøk", "Resultat", "Feilmelding", "Antall rader", "Varighet (sek)")
        sh.Rows(1).Font.Bold = True
    End If

    ' Very Hidden (kan kun hentes fram igjen via VBA-editoren, ikke via
    ' vanlig høyreklikk->Vis på en arkfane).
    ' Satt ubetinget (ikke bare ved fersk oppretting) slik at et Log-ark som
    ' allerede ble laget synlig/skjult av en eldre versjon også skjules
    ' skikkelig ved oppdatering.
    If sh.Visible <> xlSheetVeryHidden Then sh.Visible = xlSheetVeryHidden

    Set GetLogSheet = sh
End Function

' ---- Små hjelpefunksjoner ----

Public Function SheetExists(ByVal SheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Private Function IsValidSheetName(ByVal SheetNameToCheck As String) As Boolean
    Dim invalidChars As String
    Dim i As Long

    invalidChars = ":\/?*[]"
    If Len(SheetNameToCheck) = 0 Or Len(SheetNameToCheck) > 31 Then
        IsValidSheetName = False
        Exit Function
    End If
    For i = 1 To Len(invalidChars)
        If InStr(SheetNameToCheck, Mid$(invalidChars, i, 1)) > 0 Then
            IsValidSheetName = False
            Exit Function
        End If
    Next i
    IsValidSheetName = True
End Function

' Excel-tabellnavn tillater ikke mellomrom/spesialtegn slik arknavn gjor -
' bygger et trygt tabellnavn av det brukervalgte arknavnet.
Private Function SanitizeTableName(ByVal RawName As String) As String
    Dim s As String
    Dim i As Long
    Dim ch As String
    Dim result As String

    s = RawName
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = "_" Then
            result = result & ch
        Else
            result = result & "_"
        End If
    Next i
    If Len(result) = 0 Then result = "Skjema"
    If result Like "[0-9]*" Then result = "_" & result
    SanitizeTableName = result
End Function
'@

# ---- frmNettskjemaKoblinger: koblingsvindu (listen over tilkoblinger + handlinger) ----

$formCode = @'
Option Explicit

' Alle modulnivå-deklarasjoner MÅ stå her, foran første Sub/Function - VBA gir
' ellers kompileringsfeilen "Only comments may appear after End Sub...".
Private Const AVKRYSNING_BREDDE As Single = 24

Private mRegistryRows() As Long      ' listeindeks (1-basert) -> radnummer i det skjulte registeret
Private valgt() As Boolean           ' listeindeks (0-basert) -> avkrysset?
Private mRadAntall As Long
Private sisteReferanseIndeks As Long ' forrige rad klikket i avkrysningsboksen (utgangspunkt for Shift-utvalg)

Private Sub UserForm_Initialize()
    Me.Caption = "Nettskjema Henter"

    lblTittel.Font.Size = 15
    lblTittel.Font.Bold = True

    StilKnapp cmdOppdaterMarkerte, RGB(40, 150, 90), RGB(255, 255, 255), 11
    StilKnapp cmdNyTilkobling, RGB(0, 120, 215), RGB(255, 255, 255), 9
    StilKnapp cmdGaTilArk, RGB(0, 150, 160), RGB(255, 255, 255), 9
    StilKnapp cmdFjernTilkobling, RGB(200, 55, 55), RGB(255, 255, 255), 9
    StilKnapp cmdInnstillinger, RGB(112, 84, 168), RGB(255, 255, 255), 9

    ' Samme markeringsmønster som Mail-utsender: en avkrysningskolonne (kolonne
    ' 0) i selve listen. MouseUp (IKKE Click - Click fyrer aldri for denne
    ' kontrolltypen i dette miljøet) huker av/på raden når klikket treffer
    ' avkrysningsboksen, og Shift + klikk huker av/på hele intervallet siden
    ' forrige klikk.
    lstKoblinger.MultiSelect = fmMultiSelectExtended
    sisteReferanseIndeks = -1
    RefreshList
End Sub

Private Sub StilKnapp(ByVal Knapp As MSForms.CommandButton, ByVal Bakgrunn As Long, ByVal Tekst As Long, ByVal Skriftstorrelse As Single)
    Knapp.BackColor = Bakgrunn
    Knapp.ForeColor = Tekst
    Knapp.Font.Bold = True
    Knapp.Font.Size = Skriftstorrelse
End Sub

Public Sub RefreshList()
    Dim reg As Worksheet
    Dim lastRow As Long, r As Long, n As Long, i As Long
    Dim behold As Object
    Dim visning() As Variant

    ' Husk hvilke rader som var avkrysset (per registerrad), slik at
    ' avkryssingen overlever en oppdatering av selve listen.
    Set behold = CreateObject("Scripting.Dictionary")
    For i = 1 To mRadAntall
        If valgt(i - 1) Then behold(mRegistryRows(i)) = True
    Next i

    Set reg = modNettskjema.GetRegistrySheet()
    lastRow = reg.Cells(reg.Rows.Count, 1).End(xlUp).Row
    n = lastRow - 1

    sisteReferanseIndeks = -1

    If n < 1 Then
        lstKoblinger.Clear
        mRadAntall = 0
        ReDim valgt(0 To 0)
        OppdaterAntallValgt
        Exit Sub
    End If

    ReDim mRegistryRows(1 To n)
    ReDim valgt(0 To n - 1)
    ReDim visning(0 To n - 1, 0 To 4)

    For r = 2 To lastRow
        i = r - 2
        mRegistryRows(i + 1) = r
        valgt(i) = behold.Exists(r)

        visning(i, 0) = IIf(valgt(i), ChrW(9745), ChrW(9744))
        visning(i, 1) = CStr(reg.Cells(r, 1).Value)
        visning(i, 2) = CStr(reg.Cells(r, 2).Value)
        If IsDate(reg.Cells(r, 5).Value) Then
            visning(i, 3) = Format(reg.Cells(r, 5).Value, "dd.mm.yyyy hh:mm")
        Else
            visning(i, 3) = "aldri hentet"
        End If
        visning(i, 4) = CStr(reg.Cells(r, 6).Value)
    Next r

    mRadAntall = n
    lstKoblinger.ColumnCount = 5
    lstKoblinger.ColumnWidths = "24 pt;64 pt;190 pt;140 pt;60 pt"
    lstKoblinger.List = visning
    OppdaterAntallValgt
End Sub

' ---- Avkrysning (samme logikk som Mail-utsender) ----

Private Sub lstKoblinger_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim indeks As Long
    indeks = lstKoblinger.ListIndex
    If indeks = -1 Or mRadAntall = 0 Then Exit Sub

    If X < AVKRYSNING_BREDDE Then
        If ((Shift And 1) = 1) And sisteReferanseIndeks <> -1 Then
            Dim fraI As Long, tilI As Long, k As Long, tmp As Long
            fraI = sisteReferanseIndeks
            tilI = indeks
            If fraI > tilI Then
                tmp = fraI: fraI = tilI: tilI = tmp
            End If
            For k = fraI To tilI
                valgt(k) = Not valgt(k)
                lstKoblinger.List(k, 0) = IIf(valgt(k), ChrW(9745), ChrW(9744))
            Next k
        Else
            valgt(indeks) = Not valgt(indeks)
            lstKoblinger.List(indeks, 0) = IIf(valgt(indeks), ChrW(9745), ChrW(9744))
        End If
        sisteReferanseIndeks = indeks
        OppdaterAntallValgt
        Exit Sub
    End If

    sisteReferanseIndeks = indeks
End Sub

Private Sub OppdaterAntallValgt()
    Dim i As Long, n As Long
    For i = 1 To mRadAntall
        If valgt(i - 1) Then n = n + 1
    Next i
    If n > 0 Then
        cmdOppdaterMarkerte.Caption = "Oppdater markerte nettskjema (" & n & ")"
    Else
        cmdOppdaterMarkerte.Caption = "Oppdater markerte nettskjema"
    End If
End Sub

' Registerradene for alle AVKRYSSEDE skjema, i listerekkefølge (stigende radnummer).
Private Function ValgteRegistryRader() As Collection
    Dim res As Collection
    Dim i As Long
    Set res = New Collection
    For i = 1 To mRadAntall
        If valgt(i - 1) Then res.Add mRegistryRows(i)
    Next i
    Set ValgteRegistryRader = res
End Function

' Avkryssede rader - eller, hvis ingenting er avkrysset, den uthevede raden
' (brukes av "Gå til ark" og "Fjern tilkobling", som ofte gjelder ett skjema).
Private Function MaalRader() As Collection
    Dim res As Collection
    Set res = ValgteRegistryRader()
    If res.Count = 0 And mRadAntall > 0 Then
        If lstKoblinger.ListIndex >= 0 Then res.Add mRegistryRows(lstKoblinger.ListIndex + 1)
    End If
    Set MaalRader = res
End Function

' ---- Knapper ----

Private Sub cmdOppdaterMarkerte_Click()
    Dim rader As Collection
    Set rader = ValgteRegistryRader()
    If rader.Count = 0 Then
        MsgBox "Huk av ett eller flere skjema i listen først.", vbExclamation, "Oppdater markerte nettskjema"
        Exit Sub
    End If

    modNettskjema.RefreshConnectionRows rader
    RefreshList
End Sub

Private Sub cmdNyTilkobling_Click()
    modNettskjema.LeggTilTilkobling
    RefreshList
End Sub

Private Sub cmdGaTilArk_Click()
    Dim maal As Collection
    Dim sheetName As String

    Set maal = MaalRader()
    If maal.Count <> 1 Then
        MsgBox "Velg ETT skjema (huk av eller klikk på raden) for å gå til arket.", vbExclamation, "Gå til ark"
        Exit Sub
    End If

    sheetName = CStr(modNettskjema.GetRegistrySheet().Cells(CLng(maal(1)), 2).Value)
    If Not modNettskjema.SheetExists(sheetName) Then
        MsgBox "Finner ikke arket """ & sheetName & """ lenger.", vbExclamation, "Gå til ark"
        Exit Sub
    End If

    ThisWorkbook.Worksheets(sheetName).Activate
    Unload Me
End Sub

Private Sub cmdFjernTilkobling_Click()
    Dim maal As Collection
    Dim reg As Worksheet
    Dim oversikt As String
    Dim k As Long

    Set maal = MaalRader()
    If maal.Count = 0 Then
        MsgBox "Velg minst ett skjema (huk av eller klikk på raden) først.", vbExclamation, "Fjern tilkobling"
        Exit Sub
    End If

    Set reg = modNettskjema.GetRegistrySheet()
    For k = 1 To maal.Count
        oversikt = oversikt & "  - " & reg.Cells(CLng(maal(k)), 2).Value & " (Form ID " & reg.Cells(CLng(maal(k)), 1).Value & ")" & vbCrLf
    Next k

    If MsgBox("Fjerne " & IIf(maal.Count = 1, "denne tilkoblingen", "disse " & maal.Count & " tilkoblingene") & "?" & vbCrLf & vbCrLf & _
        oversikt & vbCrLf & _
        "Dette SLETTER " & IIf(maal.Count = 1, "arket", "arkene") & " og all data i " & IIf(maal.Count = 1, "det", "dem") & ". Kan ikke angres.", _
        vbExclamation + vbYesNo, "Fjern tilkobling") <> vbYes Then Exit Sub

    ' Fra bunnen og opp: RemoveConnection sletter registerraden, så radene
    ' under ville ellers ha forskjøvet radnumrene til de som står igjen.
    For k = maal.Count To 1 Step -1
        modNettskjema.RemoveConnection CLng(maal(k))
    Next k

    mRadAntall = 0 ' radnumrene er forskjøvet - ikke ta med gammel avkryssing
    RefreshList
End Sub

Private Sub cmdInnstillinger_Click()
    modNettskjema.AapneGenerelleInnstillinger
    RefreshList
End Sub
'@

# ---- frmNettskjemaInnstillinger: innstillinger (Nettskjema API-tilgang + kolonner per skjema) ----

$settingsFormCode = @'
Option Explicit

' Alle modulnivå-deklarasjoner MÅ stå her, foran første Sub/Function.
Private mFormIds() As String   ' listeindeks (1-basert) -> Form ID
Private mAntall As Long

Private Sub UserForm_Initialize()
    Me.Caption = "Nettskjema Henter - innstillinger"

    lblTittel.Font.Size = 15
    lblTittel.Font.Bold = True

    StilKnapp cmdLagre, RGB(40, 150, 90), RGB(255, 255, 255)
    StilKnapp cmdTest, RGB(0, 120, 215), RGB(255, 255, 255)
    StilKnapp cmdKolonner, RGB(112, 84, 168), RGB(255, 255, 255)
    StilKnapp cmdLukk, RGB(110, 118, 125), RGB(255, 255, 255)

    txtClientId.Text = modNettskjema.GetGlobalClientId()
    txtClientSecret.Text = modNettskjema.GetGlobalClientSecret()
    txtTokenUrl.Text = modAuthentication.TOKEN_URL
    txtApiUrl.Text = modAuthentication.API_BASE_URL

    FyllSkjemaListe
End Sub

Private Sub StilKnapp(ByVal Knapp As MSForms.CommandButton, ByVal Bakgrunn As Long, ByVal Tekst As Long)
    Knapp.BackColor = Bakgrunn
    Knapp.ForeColor = Tekst
    Knapp.Font.Bold = True
End Sub

Private Sub FyllSkjemaListe()
    Dim reg As Worksheet
    Dim lastRow As Long, r As Long

    lstSkjema.Clear
    mAntall = 0

    Set reg = modNettskjema.GetRegistrySheet()
    lastRow = reg.Cells(reg.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ReDim mFormIds(1 To lastRow - 1)
    For r = 2 To lastRow
        mAntall = mAntall + 1
        mFormIds(mAntall) = CStr(reg.Cells(r, 1).Value)
        lstSkjema.AddItem "Form " & reg.Cells(r, 1).Value & "   -   " & reg.Cells(r, 2).Value
    Next r

    lstSkjema.ListIndex = 0 ' første skjema forhåndsvalgt, så "Kolonner..." virker med en gang
End Sub

' Lagrer Client ID + Client Secret i arbeidsboken. Returnerer False (uten å
' lagre) hvis Client ID mangler.
Private Function LagreApiInnstillinger() As Boolean
    Dim id As String, hemmelighet As String

    id = Trim$(txtClientId.Text)
    hemmelighet = Trim$(txtClientSecret.Text)

    If id = "" Then
        MsgBox "Client ID kan ikke være tom.", vbExclamation, "Nettskjema API"
        Exit Function
    End If

    modNettskjema.SaveGlobalApiSettings id, hemmelighet

    If hemmelighet = "" Then
        lblStatus.ForeColor = RGB(190, 50, 50)
        lblStatus.Caption = "Lagret, men Client Secret er tom - henting vil ikke fungere."
    Else
        lblStatus.ForeColor = RGB(30, 130, 70)
        lblStatus.Caption = "Lagret i arbeidsboken kl. " & Format(Now, "hh:mm:ss") & "."
    End If
    LagreApiInnstillinger = True
End Function

Private Sub cmdLagre_Click()
    LagreApiInnstillinger
End Sub

Private Sub cmdTest_Click()
    Dim id As String, hemmelighet As String

    id = Trim$(txtClientId.Text)
    hemmelighet = Trim$(txtClientSecret.Text)
    If id = "" Then
        MsgBox "Fyll inn Client ID først.", vbExclamation, "Test tilkobling"
        Exit Sub
    End If

    ' Testen bruker de LAGREDE verdiene - lagre først hvis feltene er endret.
    If id <> modNettskjema.GetGlobalClientId() Or hemmelighet <> modNettskjema.GetGlobalClientSecret() Then
        If Not LagreApiInnstillinger() Then Exit Sub
    End If

    modNettskjema.TestTilkobling id
End Sub

Private Sub cmdKolonner_Click()
    If mAntall = 0 Then
        MsgBox "Ingen skjema er koblet til ennå. Bruk ""+ Ny tilkobling..."" i hovedvinduet først.", vbInformation, "Kolonner"
        Exit Sub
    End If
    If lstSkjema.ListIndex < 0 Then
        MsgBox "Velg et skjema i listen først.", vbExclamation, "Kolonner"
        Exit Sub
    End If

    ' Vises modalt oppå dette vinduet - når kolonnevinduet lukkes, er vi
    ' tilbake her (og hovedvinduet under er fortsatt åpent).
    modNettskjema.AapneKolonneinnstillinger mFormIds(lstSkjema.ListIndex + 1)
End Sub

Private Sub lstSkjema_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdKolonner_Click
End Sub

Private Sub cmdLukk_Click()
    Unload Me
End Sub
'@

# ---- frmKolonneInnstillinger: kolonneinnstillinger-dialog (avkrysning +
# alias per kolonne), dynamisk bygget per tilkobling - samme mønster som
# Kolonnevelger sin frmColumnPicker ----

$kolonneFormCode = @'
Option Explicit

Public FormId As String

Private mCheckBoxes As Collection
Private mTextBoxes As Collection
Private mRawKeys As Collection

Private Sub UserForm_Initialize()
    Me.Caption = "Kolonneinnstillinger"
End Sub

' Bygger en avkrysningsboks + et aliasfelt per registrert kolonne for denne
' tilkoblingen, dynamisk (antallet er ikke kjent på forhånd).
Public Sub LoadColumnSettings()
    Dim sh As Worksheet
    Set sh = modNettskjema.GetColumnSettingsSheet()

    Set mCheckBoxes = New Collection
    Set mTextBoxes = New Collection
    Set mRawKeys = New Collection

    lblHjelp.Caption = "Huk av for å vise en kolonne. Bruk eksempelsvaret til å se hvilket spørsmål det er, og skriv gjerne et alias."

    ' Finn tabellen for denne tilkoblingen (hvis den finnes) - brukes til å
    ' vise et EKTE eksempelsvar per kolonne, siden APIet ikke gir oss den
    ' ekte spørsmålsteksten (se feedback-vba-com-test-gotchas i minnet -
    ' /elements ga 401 for denne API-klienten).
    Dim tbl As ListObject
    Dim hasTable As Boolean
    hasTable = False
    Dim reg As Worksheet
    Set reg = modNettskjema.GetRegistrySheet()
    Dim regRow As Long
    regRow = modNettskjema.FindConnectionByFormId(FormId)
    If regRow > 0 Then
        Dim sheetNameHere As String, tableNameHere As String
        sheetNameHere = CStr(reg.Cells(regRow, 2).Value)
        tableNameHere = CStr(reg.Cells(regRow, 3).Value)
        On Error Resume Next
        Set tbl = ThisWorkbook.Worksheets(sheetNameHere).ListObjects(tableNameHere)
        On Error GoTo 0
        hasTable = Not tbl Is Nothing
    End If

    Dim lastRow As Long, r As Long
    lastRow = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row

    Dim rowHeight As Single
    rowHeight = 40
    Dim idx As Long
    idx = 0

    For r = 2 To lastRow
        If CStr(sh.Cells(r, 1).Value) = FormId Then
            idx = idx + 1

            Dim rawKeyHere As String
            rawKeyHere = CStr(sh.Cells(r, 2).Value)
            Dim aliasHere As String
            aliasHere = CStr(sh.Cells(r, 4).Value)

            Dim chk As MSForms.CheckBox
            Set chk = frameColumns.Controls.Add("Forms.CheckBox.1", "chkVis" & idx, True)
            chk.Left = 6
            chk.Top = 10 + (idx - 1) * rowHeight
            chk.Width = 20
            chk.Value = (sh.Cells(r, 3).Value <> False)

            Dim exampleText As String
            exampleText = GetExampleAnswer(tbl, hasTable, IIf(aliasHere <> "", aliasHere, rawKeyHere))

            Dim lbl As MSForms.Label
            Set lbl = frameColumns.Controls.Add("Forms.Label.1", "lblKey" & idx, True)
            lbl.Left = 28
            lbl.Top = 4 + (idx - 1) * rowHeight
            lbl.Width = 260
            lbl.Height = 32
            lbl.WordWrap = True
            If exampleText <> "" Then
                lbl.Caption = rawKeyHere & vbCrLf & "F.eks.: " & exampleText
            Else
                lbl.Caption = rawKeyHere & vbCrLf & "(ingen svar å vise enda)"
            End If

            Dim txt As MSForms.TextBox
            Set txt = frameColumns.Controls.Add("Forms.TextBox.1", "txtAlias" & idx, True)
            txt.Left = 294
            txt.Top = 10 + (idx - 1) * rowHeight
            txt.Width = 210
            txt.Height = 20
            txt.Text = aliasHere

            mCheckBoxes.Add chk
            mTextBoxes.Add txt
            mRawKeys.Add rawKeyHere
        End If
    Next r

    If idx = 0 Then
        Dim lblIngen As MSForms.Label
        Set lblIngen = frameColumns.Controls.Add("Forms.Label.1", "lblIngen", True)
        lblIngen.Left = 6
        lblIngen.Top = 6
        lblIngen.Width = 480
        lblIngen.Height = 18
        lblIngen.Caption = "Ingen kolonner registrert enda - hent svar minst en gang først."
    End If

    Dim neededHeight As Single
    neededHeight = 6 + idx * rowHeight + 10
    If neededHeight > frameColumns.Height Then
        frameColumns.ScrollBars = fmScrollBarsVertical
        frameColumns.ScrollHeight = neededHeight
    Else
        frameColumns.ScrollBars = fmScrollBarsNone
    End If
End Sub

' Henter første ikke-tomme svar i den gitte kolonnen (identifisert ved sitt
' NAAVAERENDE visningsnavn - RawKey eller alias) fra allerede hentet data, som
' et hint om hvilket spørsmål RawKey egentlig er - siden APIet selv ikke gir
' oss lesbar spørsmålstekst. Trunkeres for å holde etiketten kompakt.
Private Function GetExampleAnswer(ByVal Tbl As ListObject, ByVal HasTable As Boolean, ByVal CurrentDisplayName As String) As String
    GetExampleAnswer = ""
    If Not HasTable Then Exit Function
    If Tbl.ListRows.Count = 0 Then Exit Function

    On Error Resume Next
    Dim col As ListColumn
    Set col = Tbl.ListColumns(CurrentDisplayName)
    On Error GoTo 0
    If col Is Nothing Then Exit Function

    Dim dataRow As Long
    Dim cellVal As String
    For dataRow = 1 To col.DataBodyRange.Rows.Count
        cellVal = Trim$(CStr(col.DataBodyRange.Cells(dataRow, 1).Value))
        If cellVal <> "" Then
            If Len(cellVal) > 40 Then cellVal = Left$(cellVal, 40) & "..."
            GetExampleAnswer = cellVal
            Exit Function
        End If
    Next dataRow
End Function

Private Sub cmdLagre_Click()
    Dim sh As Worksheet
    Set sh = modNettskjema.GetColumnSettingsSheet()

    Dim reg As Worksheet
    Set reg = modNettskjema.GetRegistrySheet()
    Dim regRow As Long
    regRow = modNettskjema.FindConnectionByFormId(FormId)

    Dim tbl As ListObject
    Dim hasTable As Boolean
    hasTable = False
    If regRow > 0 Then
        Dim sheetName As String, tableName As String
        sheetName = CStr(reg.Cells(regRow, 2).Value)
        tableName = CStr(reg.Cells(regRow, 3).Value)
        On Error Resume Next
        Set tbl = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
        On Error GoTo 0
        hasTable = Not tbl Is Nothing
    End If

    Dim i As Long
    For i = 1 To mRawKeys.Count
        Dim rawKey As String
        rawKey = mRawKeys(i)
        Dim r As Long
        r = modNettskjema.FindColumnSettingsRowPublic(FormId, rawKey)
        If r > 0 Then
            Dim oldAlias As String
            oldAlias = CStr(sh.Cells(r, 4).Value)
            Dim oldDisplayName As String
            oldDisplayName = IIf(oldAlias <> "", oldAlias, rawKey)

            Dim newAlias As String
            newAlias = Trim$(mTextBoxes(i).Text)
            Dim newDisplayName As String
            newDisplayName = IIf(newAlias <> "", newAlias, rawKey)

            sh.Cells(r, 3).Value = mCheckBoxes(i).Value
            sh.Cells(r, 4).Value = newAlias

            If hasTable Then
                On Error Resume Next
                Dim col As ListColumn
                Set col = tbl.ListColumns(oldDisplayName)
                If Not col Is Nothing Then
                    If oldDisplayName <> newDisplayName Then col.Name = newDisplayName
                    col.Range.EntireColumn.Hidden = Not mCheckBoxes(i).Value
                End If
                Set col = Nothing
                On Error GoTo 0
            End If
        End If
    Next i

    Unload Me
    MsgBox "Kolonneinnstillinger lagret og brukt.", vbInformation
End Sub

Private Sub cmdAvbryt_Click()
    Unload Me
End Sub
'@

function Get-NettskjemaHenterVersion {
    param([string]$Code)
    if ($Code -match 'NETTSKJEMA_HENTER_VERSION\s*As\s*String\s*=\s*"([^"]+)"') {
        return $Matches[1]
    }
    return $null
}

$sourceVersion = Get-NettskjemaHenterVersion $mainCode

$allComponents = @(
    @{ Name = "JsonConverter";          Code = $jsonConverterCode; File = "JsonConverter.bas"; Type = 1 },
    @{ Name = "modCredentialManager";   Code = $credMgrCode;       File = "modCredentialManager.bas"; Type = 1 },
    @{ Name = "modAuthentication";      Code = $authCode;          File = "modAuthentication.bas"; Type = 1 },
    @{ Name = "modJsonFlatten";         Code = $flattenCode;       File = "modJsonFlatten.bas"; Type = 1 },
    @{ Name = "modNettskjema";          Code = $mainCode;          File = "modNettskjema.bas"; Type = 1 },
    @{ Name = "frmNettskjemaKoblinger"; Code = $formCode;          File = "frmNettskjemaKoblinger.frm"; Type = 3 },
    @{ Name = "frmKolonneInnstillinger"; Code = $kolonneFormCode;  File = "frmKolonneInnstillinger.frm"; Type = 4 },
    @{ Name = "frmNettskjemaInnstillinger"; Code = $settingsFormCode; File = "frmNettskjemaInnstillinger.frm"; Type = 5 }
)

$buttonOnActions = @("AapneNettskjemaHenter")

# ---- 1. Finn målfilen ----

if (-not $Path) {
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Excel-filer (*.xlsm)|*.xlsm|Alle filer (*.*)|*.*"
    $dlg.Title = "Velg Excel-fil som skal få Nettskjema Henter (opprett en tom .xlsm først om nødvendig)"
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Output "Avbrutt."
        exit 0
    }
    $Path = $dlg.FileName
}

if (-not (Test-Path $Path)) {
    Write-Output "FEIL: Fant ikke filen: $Path"
    exit 1
}
$Path = (Resolve-Path $Path).Path

# ---- 2. Finn en åpen Excel-instans som har filen, ellers åpne den selv ----

$excel = $null
$wb = $null

try {
    if ($NyExcelInstans) {
        $running = $null
    } else {
        $running = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    }
    if ($running) {
        foreach ($w in $running.Workbooks) {
            if ($w.FullName -eq $Path) { $wb = $w; $excel = $running; break }
        }
        if ($null -eq $excel) { $excel = $running }
    }
} catch {
    $excel = $null
}

if ($null -eq $excel) {
    $excel = New-Object -ComObject Excel.Application
}
$excel.Visible = $true
$excel.DisplayAlerts = $false

if ($null -eq $wb) {
    $wb = $excel.Workbooks.Open($Path)
}
$wb.Activate()

try {
    $vbproj = $wb.VBProject
    $null = $vbproj.VBComponents.Count
} catch {
    Write-Output "FEIL: Får ikke tilgang til VBA-prosjektet."
    Write-Output "Sjekk at 'Klarer tilgang til VBA-prosjektobjektmodellen' er huket av i Excel sitt Klareringssenter (Filer > Alternativer > Klareringssenter > Innstillinger for klareringssenteret > Makroinnstillinger)."
    exit 1
}

if ($Uninstall) {
    # Fjerner KUN denne makroens egne, eksakt navngitte VBA-komponenter og
    # knappene med denne makroens EGNE OnAction-navn - aldri et generisk
    # filter (et ark kan ha flere uavhengig installerte verktøys knapper
    # side om side). Dashboard/Settings/Data/Log-arkene og innholdet i dem
    # (inkl. hentede submissions) røres IKKE - kun kode + knapper fjernes.
    Write-Output "Fjerner Nettskjema Henter fra $($wb.Name) ..."
    foreach ($item in $allComponents) {
        try {
            $eksisterende = $vbproj.VBComponents.Item($item.Name)
            $vbproj.VBComponents.Remove($eksisterende)
            Write-Output "  Fjernet $($item.Name)"
        } catch {}
    }
    try {
        foreach ($ws in $wb.Worksheets) {
            foreach ($btn in @($ws.Buttons())) {
                if ($buttonOnActions -contains $btn.OnAction) { $btn.Delete() }
            }
        }
    } catch {}
    try { $wb.Save() } catch {
        Write-Output "FEIL ved lagring: $($_.Exception.Message)"
        exit 1
    }
    Write-Output "Ferdig - Nettskjema Henter er fjernet fra $($wb.Name). Ark og hentet data er beholdt."
    exit 0
}

# ---- 3. Sjekk hva som allerede er installert ----

$existingModule = $null
foreach ($comp in $vbproj.VBComponents) {
    if ($comp.Name -eq "modNettskjema") { $existingModule = $comp }
}

$installedVersion = $null
if ($existingModule) {
    # CountOfLines kan være 0 hvis komponenten finnes men er tom (sett en gang
    # etter en Excel-krasj midt i en installasjon) - Lines(1, 0) er ugyldig og
    # kaster "Invalid procedure call or argument". Behandle det som "ingen
    # gyldig versjon funnet" (samme som om komponenten ikke fantes) i stedet
    # for å krasje - fresh-install-logikken under håndterer det trygt
    # (fjerner den tomme komponenten og importerer en frisk versjon).
    if ($existingModule.CodeModule.CountOfLines -gt 0) {
        $existingCode = $existingModule.CodeModule.Lines(1, $existingModule.CodeModule.CountOfLines)
        $installedVersion = Get-NettskjemaHenterVersion $existingCode
    } else {
        Write-Output "ADVARSEL: fant modNettskjema, men den er tom (0 linjer) - sannsynligvis fra en tidligere avbrutt installasjon. Installerer på nytt (inkl. knapp)."
        $existingModule = $null
    }
}

$isFreshInstall = ($null -eq $existingModule)

if (-not $isFreshInstall -and $installedVersion -eq $sourceVersion -and -not $Force) {
    Write-Output "Nettskjema Henter er allerede installert og oppdatert (versjon $installedVersion) i $($wb.Name)."
    exit 0
}

if ($isFreshInstall) {
    Write-Output "Installerer Nettskjema Henter (versjon $sourceVersion) i $($wb.Name) ..."
} else {
    Write-Output "Oppdaterer Nettskjema Henter: versjon $installedVersion -> $sourceVersion i $($wb.Name) ..."
}

# ---- 4. Bygg de seks komponentene i en midlertidig arbeidsbok, eksporter til en temp-mappe ----

$originalActiveSheetName = $wb.ActiveSheet.Name
$tempDir = Join-Path $env:TEMP ("NettskjemaHenter_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempDir | Out-Null

try {
    $buildWb = $excel.Workbooks.Add()
    $buildProj = $buildWb.VBProject

    foreach ($item in $allComponents) {
        if ($item.Type -eq 3) {
            # UserForm (Designer) - frmNettskjemaKoblinger: hovedmeny. Farger, fet
            # skrift og avkrysningsmønster settes i VBA (UserForm_Initialize).
            $comp = $buildProj.VBComponents.Add(3)
            $comp.Name = $item.Name
            $comp.Properties("Width").Value = 540
            $comp.Properties("Height").Value = 394
            $comp.Properties("Caption").Value = "Nettskjema Henter"
            $designer = $comp.Designer

            $lblTittel = $designer.Controls.Add("Forms.Label.1", "lblTittel", $true)
            $lblTittel.Left = 12; $lblTittel.Top = 8; $lblTittel.Width = 260; $lblTittel.Height = 24
            $lblTittel.Caption = "Nettskjema Henter"

            $lblUndertittel = $designer.Controls.Add("Forms.Label.1", "lblUndertittel", $true)
            $lblUndertittel.Left = 12; $lblUndertittel.Top = 33; $lblUndertittel.Width = 380; $lblUndertittel.Height = 14
            $lblUndertittel.Caption = "Huk av ett eller flere skjema i listen, og oppdater dem samlet."
            $lblUndertittel.ForeColor = 6579300

            $cmdInnstillinger = $designer.Controls.Add("Forms.CommandButton.1", "cmdInnstillinger", $true)
            $cmdInnstillinger.Left = 400; $cmdInnstillinger.Top = 10; $cmdInnstillinger.Width = 116; $cmdInnstillinger.Height = 28
            $cmdInnstillinger.Caption = "Innstillinger..."

            # Kolonneoverskrifter over listen (ListBox.List har ingen egne overskrifter).
            $lblKolForm = $designer.Controls.Add("Forms.Label.1", "lblKolForm", $true)
            $lblKolForm.Left = 45; $lblKolForm.Top = 53; $lblKolForm.Width = 60; $lblKolForm.Height = 13
            $lblKolForm.Caption = "Form ID"
            $lblKolForm.ForeColor = 6579300

            $lblKolArk = $designer.Controls.Add("Forms.Label.1", "lblKolArk", $true)
            $lblKolArk.Left = 109; $lblKolArk.Top = 53; $lblKolArk.Width = 180; $lblKolArk.Height = 13
            $lblKolArk.Caption = "Ark"
            $lblKolArk.ForeColor = 6579300

            $lblKolSist = $designer.Controls.Add("Forms.Label.1", "lblKolSist", $true)
            $lblKolSist.Left = 299; $lblKolSist.Top = 53; $lblKolSist.Width = 130; $lblKolSist.Height = 13
            $lblKolSist.Caption = "Sist hentet"
            $lblKolSist.ForeColor = 6579300

            $lblKolSvar = $designer.Controls.Add("Forms.Label.1", "lblKolSvar", $true)
            $lblKolSvar.Left = 439; $lblKolSvar.Top = 53; $lblKolSvar.Width = 60; $lblKolSvar.Height = 13
            $lblKolSvar.Caption = "Svar"
            $lblKolSvar.ForeColor = 6579300

            $lstKoblinger = $designer.Controls.Add("Forms.ListBox.1", "lstKoblinger", $true)
            $lstKoblinger.Left = 12; $lstKoblinger.Top = 67; $lstKoblinger.Width = 504; $lstKoblinger.Height = 162

            $cmdOppdaterMarkerte = $designer.Controls.Add("Forms.CommandButton.1", "cmdOppdaterMarkerte", $true)
            $cmdOppdaterMarkerte.Left = 12; $cmdOppdaterMarkerte.Top = 240; $cmdOppdaterMarkerte.Width = 504; $cmdOppdaterMarkerte.Height = 40
            $cmdOppdaterMarkerte.Caption = "Oppdater markerte nettskjema"

            $fraTilkoblinger = $designer.Controls.Add("Forms.Frame.1", "fraTilkoblinger", $true)
            $fraTilkoblinger.Left = 12; $fraTilkoblinger.Top = 292; $fraTilkoblinger.Width = 504; $fraTilkoblinger.Height = 60
            $fraTilkoblinger.Caption = "Tilkoblinger"

            $cmdNyTilkobling = $fraTilkoblinger.Controls.Add("Forms.CommandButton.1", "cmdNyTilkobling", $true)
            $cmdNyTilkobling.Left = 8; $cmdNyTilkobling.Top = 12; $cmdNyTilkobling.Width = 156; $cmdNyTilkobling.Height = 30
            $cmdNyTilkobling.Caption = "+ Ny tilkobling..."

            $cmdGaTilArk = $fraTilkoblinger.Controls.Add("Forms.CommandButton.1", "cmdGaTilArk", $true)
            $cmdGaTilArk.Left = 172; $cmdGaTilArk.Top = 12; $cmdGaTilArk.Width = 156; $cmdGaTilArk.Height = 30
            $cmdGaTilArk.Caption = "Gå til ark"

            $cmdFjernTilkobling = $fraTilkoblinger.Controls.Add("Forms.CommandButton.1", "cmdFjernTilkobling", $true)
            $cmdFjernTilkobling.Left = 336; $cmdFjernTilkobling.Top = 12; $cmdFjernTilkobling.Width = 156; $cmdFjernTilkobling.Height = 30
            $cmdFjernTilkobling.Caption = "Fjern tilkobling"

            $comp.CodeModule.AddFromString($item.Code)
            $comp.Export((Join-Path $tempDir $item.File))
        } elseif ($item.Type -eq 4) {
            # UserForm (Designer) - frmKolonneInnstillinger: tittel + rulleoar
            # ramme (fylles dynamisk med avkrysning+alias-felt av VBA-koden
            # selv, ett sett per kolonne - antallet er ikke kjent her) + to
            # knapper.
            $comp = $buildProj.VBComponents.Add(3)
            $comp.Name = $item.Name
            $comp.Properties("Width").Value = 540
            $comp.Properties("Height").Value = 400
            $comp.Properties("Caption").Value = "Kolonneinnstillinger"
            $designer = $comp.Designer

            $lblHjelp = $designer.Controls.Add("Forms.Label.1", "lblHjelp", $true)
            $lblHjelp.Left = 10; $lblHjelp.Top = 8; $lblHjelp.Width = 510; $lblHjelp.Height = 28
            $lblHjelp.Caption = "Huk av for å vise en kolonne. Skriv et alias for å gi spørsmålet et lesbart navn (valgfritt)."

            $frameColumns = $designer.Controls.Add("Forms.Frame.1", "frameColumns", $true)
            $frameColumns.Left = 10; $frameColumns.Top = 40; $frameColumns.Width = 510; $frameColumns.Height = 300
            $frameColumns.Caption = ""
            $frameColumns.ScrollBars = 2

            $cmdLagre = $designer.Controls.Add("Forms.CommandButton.1", "cmdLagre", $true)
            $cmdLagre.Left = 10; $cmdLagre.Top = 350; $cmdLagre.Width = 250; $cmdLagre.Height = 30
            $cmdLagre.Caption = "Lagre og bruk"

            $cmdAvbryt = $designer.Controls.Add("Forms.CommandButton.1", "cmdAvbryt", $true)
            $cmdAvbryt.Left = 270; $cmdAvbryt.Top = 350; $cmdAvbryt.Width = 250; $cmdAvbryt.Height = 30
            $cmdAvbryt.Caption = "Avbryt"

            $comp.CodeModule.AddFromString($item.Code)
            $comp.Export((Join-Path $tempDir $item.File))
        } elseif ($item.Type -eq 5) {
            # UserForm (Designer) - frmNettskjemaInnstillinger: API-tilgang
            # (Client ID/Secret, faste adresser, test) + kolonneinnstillinger
            # per skjema. Farger/fet skrift settes i VBA (UserForm_Initialize).
            $comp = $buildProj.VBComponents.Add(3)
            $comp.Name = $item.Name
            $comp.Properties("Width").Value = 500
            $comp.Properties("Height").Value = 490
            $comp.Properties("Caption").Value = "Nettskjema Henter - innstillinger"
            $designer = $comp.Designer

            $lblTittel = $designer.Controls.Add("Forms.Label.1", "lblTittel", $true)
            $lblTittel.Left = 12; $lblTittel.Top = 8; $lblTittel.Width = 300; $lblTittel.Height = 24
            $lblTittel.Caption = "Innstillinger"

            # --- Seksjon 1: Nettskjema API ---
            $fraApi = $designer.Controls.Add("Forms.Frame.1", "fraApi", $true)
            $fraApi.Left = 12; $fraApi.Top = 38; $fraApi.Width = 466; $fraApi.Height = 232
            $fraApi.Caption = "Nettskjema API"

            $lblClientId = $fraApi.Controls.Add("Forms.Label.1", "lblClientId", $true)
            $lblClientId.Left = 10; $lblClientId.Top = 18; $lblClientId.Width = 92; $lblClientId.Height = 16
            $lblClientId.Caption = "Client ID:"

            $txtClientId = $fraApi.Controls.Add("Forms.TextBox.1", "txtClientId", $true)
            $txtClientId.Left = 106; $txtClientId.Top = 15; $txtClientId.Width = 346; $txtClientId.Height = 20

            $lblClientSecret = $fraApi.Controls.Add("Forms.Label.1", "lblClientSecret", $true)
            $lblClientSecret.Left = 10; $lblClientSecret.Top = 46; $lblClientSecret.Width = 92; $lblClientSecret.Height = 16
            $lblClientSecret.Caption = "Client Secret:"

            $txtClientSecret = $fraApi.Controls.Add("Forms.TextBox.1", "txtClientSecret", $true)
            $txtClientSecret.Left = 106; $txtClientSecret.Top = 43; $txtClientSecret.Width = 346; $txtClientSecret.Height = 20

            $lblVarsel = $fraApi.Controls.Add("Forms.Label.1", "lblVarsel", $true)
            $lblVarsel.Left = 106; $lblVarsel.Top = 68; $lblVarsel.Width = 346; $lblVarsel.Height = 28
            $lblVarsel.WordWrap = $true
            $lblVarsel.Caption = "Client Secret lagres i selve arbeidsboken (i et skjult ark). Behandle filen som sensitiv - alle med tilgang til den kan i prinsippet lese secret-en."
            $lblVarsel.ForeColor = 6579300

            $lblTokenUrl = $fraApi.Controls.Add("Forms.Label.1", "lblTokenUrl", $true)
            $lblTokenUrl.Left = 10; $lblTokenUrl.Top = 106; $lblTokenUrl.Width = 92; $lblTokenUrl.Height = 16
            $lblTokenUrl.Caption = "Token-adresse:"

            $txtTokenUrl = $fraApi.Controls.Add("Forms.TextBox.1", "txtTokenUrl", $true)
            $txtTokenUrl.Left = 106; $txtTokenUrl.Top = 103; $txtTokenUrl.Width = 346; $txtTokenUrl.Height = 20
            $txtTokenUrl.Locked = $true
            $txtTokenUrl.BackColor = 15921906

            $lblApiUrl = $fraApi.Controls.Add("Forms.Label.1", "lblApiUrl", $true)
            $lblApiUrl.Left = 10; $lblApiUrl.Top = 134; $lblApiUrl.Width = 92; $lblApiUrl.Height = 16
            $lblApiUrl.Caption = "API-adresse:"

            $txtApiUrl = $fraApi.Controls.Add("Forms.TextBox.1", "txtApiUrl", $true)
            $txtApiUrl.Left = 106; $txtApiUrl.Top = 131; $txtApiUrl.Width = 346; $txtApiUrl.Height = 20
            $txtApiUrl.Locked = $true
            $txtApiUrl.BackColor = 15921906

            $cmdLagre = $fraApi.Controls.Add("Forms.CommandButton.1", "cmdLagre", $true)
            $cmdLagre.Left = 106; $cmdLagre.Top = 164; $cmdLagre.Width = 170; $cmdLagre.Height = 30
            $cmdLagre.Caption = "Lagre"

            $cmdTest = $fraApi.Controls.Add("Forms.CommandButton.1", "cmdTest", $true)
            $cmdTest.Left = 282; $cmdTest.Top = 164; $cmdTest.Width = 170; $cmdTest.Height = 30
            $cmdTest.Caption = "Test tilkobling"

            $lblStatus = $fraApi.Controls.Add("Forms.Label.1", "lblStatus", $true)
            $lblStatus.Left = 106; $lblStatus.Top = 200; $lblStatus.Width = 346; $lblStatus.Height = 16
            $lblStatus.Caption = ""

            # --- Seksjon 2: Kolonner per skjema ---
            $fraKolonner = $designer.Controls.Add("Forms.Frame.1", "fraKolonner", $true)
            $fraKolonner.Left = 12; $fraKolonner.Top = 282; $fraKolonner.Width = 466; $fraKolonner.Height = 132
            $fraKolonner.Caption = "Kolonner per skjema"

            $lblKolHjelp = $fraKolonner.Controls.Add("Forms.Label.1", "lblKolHjelp", $true)
            $lblKolHjelp.Left = 10; $lblKolHjelp.Top = 16; $lblKolHjelp.Width = 442; $lblKolHjelp.Height = 26
            $lblKolHjelp.WordWrap = $true
            $lblKolHjelp.Caption = "Velg et skjema og åpne kolonneinnstillingene: hvilke kolonner som vises, og eget navn (alias) per kolonne."
            $lblKolHjelp.ForeColor = 6579300

            $lstSkjema = $fraKolonner.Controls.Add("Forms.ListBox.1", "lstSkjema", $true)
            $lstSkjema.Left = 10; $lstSkjema.Top = 46; $lstSkjema.Width = 316; $lstSkjema.Height = 74

            $cmdKolonner = $fraKolonner.Controls.Add("Forms.CommandButton.1", "cmdKolonner", $true)
            $cmdKolonner.Left = 336; $cmdKolonner.Top = 46; $cmdKolonner.Width = 116; $cmdKolonner.Height = 34
            $cmdKolonner.Caption = "Kolonner..."

            $cmdLukk = $designer.Controls.Add("Forms.CommandButton.1", "cmdLukk", $true)
            $cmdLukk.Left = 368; $cmdLukk.Top = 424; $cmdLukk.Width = 110; $cmdLukk.Height = 28
            $cmdLukk.Caption = "Lukk"

            $comp.CodeModule.AddFromString($item.Code)
            $comp.Export((Join-Path $tempDir $item.File))
        } else {
            $comp = $buildProj.VBComponents.Add(1)   # vbext_ct_StdModule
            $comp.Name = $item.Name
            $comp.CodeModule.AddFromString($item.Code)
            $comp.Export((Join-Path $tempDir $item.File))
        }
    }

    $buildWb.Close($false)

    # ---- 5. Bytt ut / legg til komponentene i mål-arbeidsboken ----

    foreach ($item in $allComponents) {
        $filePath = Join-Path $tempDir $item.File

        $existing = $null
        foreach ($comp in $vbproj.VBComponents) {
            if ($comp.Name -eq $item.Name) { $existing = $comp }
        }
        if ($existing) {
            $vbproj.VBComponents.Remove($existing)
        }
        $null = $vbproj.VBComponents.Import($filePath)
        Write-Output "  $($item.Name) installert"
    }

    # Ingen automatisk ark-knapp lenger, verken ved fersk installasjon eller
    # oppdatering (Håkons eksplisitte onske 2026-09-22: en makro legger seg
    # bare til - en knapp/inngang i arket er et eget, manuelt valg brukeren
    # gjor selv, typisk via Makromeny). Registeret (skjult ark) og Log-arket
    # lager seg selv forste gang de trengs (se GetRegistrySheet/GetLogSheet
    # i VBA) - upåvirket av denne endringen.

    Write-Output ""
    Write-Output "Ferdig. Husk å lagre filen selv (Ctrl+S) når du er klar."
    if ($isFreshInstall) {
        Write-Output ""
        Write-Output "Neste steg:"
        Write-Output "  1. Kjør ""AapneNettskjemaHenter"" (f.eks. via Makromeny - ingen knapp legges til i arket automatisk lenger)."
        Write-Output "  2. Trykk ""+ Ny tilkobling..."" og fyll inn Form ID, Client ID og et arknavn."
        Write-Output "  3. Første gang: lim inn Client Secret når du blir spurt (opprettet på authorization.nettskjema.no)."
        Write-Output "  4. Skjemaet får sitt eget ark + tabell, og hentes automatisk med en gang."
    }
} finally {
    Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
}
