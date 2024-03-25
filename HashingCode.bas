Attribute VB_Name = "HashingCode"
Public Function SHA1(ByVal s As String) As String
    Dim Enc As Object, Prov As Object
    Dim Hash() As Byte, i As Integer

    Set Enc = CreateObject("System.Text.UTF8Encoding")
    Set Prov = CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")

    Hash = Prov.ComputeHash_2(Enc.GetBytes_4(s))

    SHA1 = ""
    For i = LBound(Hash) To UBound(Hash)
        SHA1 = SHA1 & Hex(Hash(i) \ 16) & Hex(Hash(i) Mod 16)
    Next
End Function


Sub encryptPass()


End Sub
