Attribute VB_Name = "Module6"
Option Explicit

'========================================================
' KATERN GENERATOR (Geoptimaliseerd met KaternLijst)
' Modulebreed katern-array met Range-objecten
'========================================================

Private Katern() As Range

Private Const MAX_STROKEN As Long = 8
Private Const PAGINAS_PER_STROOK As Long = 4



'========================================================
' HOOFDPROGRAMMA
'========================================================

Public Sub MaakKaternen()

    Dim BronDocument As Document
    Dim BoekPaginas As Collection
    Dim KaternMaten As Variant
    
    Dim KaternNr As Long
    Dim EerstePagina As Long
    Dim n As Long
    
    Dim DoelMap As String
    Dim TemplatePad As String

    Set BronDocument = ActiveDocument

    ' 1. Lees alle pagina's in een collectie
    Set BoekPaginas = LeesBoekPaginas(BronDocument)
    Debug.Print "Aantal pagina's: "; BoekPaginas.Count

    If BoekPaginas.Count = 0 Then Exit Sub

    ' 2. Bereken de optimale katernlijst
    KaternMaten = BerekenOptimaleKaternLijst(BoekPaginas.Count)

    DoelMap = MaakDoelMap(BronDocument)
    TemplatePad = TemplateNaam("KaternTemplat.dotx")
    
    If TemplatePad = "" Then Exit Sub

    EerstePagina = 1

    ' 3. Loop door de berekende katern-array
    For KaternNr = LBound(KaternMaten) To UBound(KaternMaten)
        
        ' Maat van dit katern in stroken/vellen
        n = KaternMaten(KaternNr)

        ' Maak het katern aan en sla het op
        MaakEenKatern _
            KaternNr, _
            EerstePagina, _
            n, _
            BoekPaginas, _
            TemplatePad, _
            DoelMap

        ' Startpagina voor het volgende katern
        EerstePagina = EerstePagina + _
                       (PAGINAS_PER_STROOK * n)
        
    Next KaternNr

    MsgBox "Klaar! Alle " & _
           (UBound(KaternMaten) - LBound(KaternMaten) + 1) & _
           " katernen zijn succesvol gegenereerd.", _
           vbInformation, _
           "Koppeling Voltooid"

End Sub



'========================================================
' OPTIMALE KATERN LIJST FUNCTIE (Array 1 tot n met maten)
'========================================================

Private Function BerekenOptimaleKaternLijst(ByVal totaalPaginas As Long) As Variant
    Dim aantalVellen As Long
    Dim aantalKaternen As Long
    Dim basisVellen As Long
    Dim restVellen As Long
    Dim katernVellen() As Long
    Dim i As Long, j As Long
    
    ' Reken om naar vellen (1 vel = 4 pagina's, rond naar boven af op een viertal)
    totaalPaginas = ((totaalPaginas + 3) \ PAGINAS_PER_STROOK) * PAGINAS_PER_STROOK
    aantalVellen = totaalPaginas / PAGINAS_PER_STROOK
    
    ' Bepaal het minimale aantal katernen op basis van max 8 vellen per katern
    aantalKaternen = (aantalVellen + MAX_STROKEN - 1) \ MAX_STROKEN
    If aantalKaternen < 1 Then aantalKaternen = 1
    
    ' Zorg dat het gemiddelde per katern niet hoger is dan het maximum
    Do While (aantalVellen / aantalKaternen > MAX_STROKEN)
        aantalKaternen = aantalKaternen + 1
    Loop
    
    ' Maak de array aan van 1 tot n katernen
    ReDim katernVellen(1 To aantalKaternen)
    
    basisVellen = aantalVellen \ aantalKaternen
    restVellen = aantalVellen Mod aantalKaternen
    
    ' Vul de array met de basismaat
    For i = 1 To aantalKaternen
        katernVellen(i) = basisVellen
    Next i
    
    ' Verdeel de restvellen over de voorgaande katernen (maximaal MAX_STROKEN per katern)
    For i = 1 To restVellen
        For j = 1 To aantalKaternen
            If katernVellen(j) < MAX_STROKEN Then
                katernVellen(j) = katernVellen(j) + 1
                Exit For
            End If
        Next j
    Next i
    
    ' Retourneer de array met maten
    BerekenOptimaleKaternLijst = katernVellen
End Function


'========================================================
' EEN KATERN
'========================================================

Private Sub MaakEenKatern( _
        ByVal KaternNr As Long, _
        ByVal EerstePagina As Long, _
        ByVal n As Long, _
        ByVal BoekPaginas As Collection, _
        ByVal TemplatePad As String, _
        ByVal DoelMap As String)

    Dim DoelDocument As Document
    Dim DoelTabel As Table
    
    Dim np As Long
    Dim Naam As String
    
' --- SCHERMVERVERSING UITZETTEN ---
    ' Hierdoor zie je geen lege tabellen of templates meer op het scherm flitsen
    Application.ScreenUpdating = False
    On Error GoTo FoutAfhandeling
    
    MaakKaternArray _
        EerstePagina, _
        n, _
        BoekPaginas

    Set DoelDocument = _
        MaakDoelDocument( _
            TemplatePad, _
            2 * n, _
            np)

    If DoelDocument Is Nothing Then Exit Sub

    Set DoelTabel = DoelDocument.Tables(1)

    KopieerKaternNaarDoel _
        DoelTabel, _
        np, _
        2 * n

    Naam = DoelMap & _
            "\Katern" & _
            Format(KaternNr, "00") & _
            ".docx"

    DoelDocument.SaveAs Naam

    Debug.Print "Opslaan: "; Naam

    DoelDocument.Close False

FoutAfhandeling:
    If Err.Number <> 0 Then
        MsgBox "Er is een fout opgetreden: " & Err.Description, vbCritical
    End If

Einde:
    ' --- SCHERMVERVERSING ALTIJD WEER AANZETTEN ---
    Application.ScreenUpdating = True

    
End Sub


'========================================================
' BOEKPAGINA'S LEZEN
'========================================================

Private Function LeesBoekPaginas( _
        ByVal Doc As Document) As Collection

    Dim Paginas As New Collection
    Dim StartPos As Long
    Dim EindPos As Long
    Dim r As Range
    Dim p As Range
    Dim Nr As Long
    Dim Pos As Long

    StartPos = Doc.Range.Start
    EindPos = Doc.Range.End
    Nr = 1

    Do While StartPos < EindPos

        Set r = Doc.Range(StartPos, EindPos)
        Pos = InStr(1, r.Text, Chr(12))

        If Pos > 0 Then
            Set p = Doc.Range(StartPos, StartPos + Pos - 1)
        Else
            Set p = Doc.Range(StartPos, EindPos)
        End If

        Paginas.Add p.Duplicate

        Debug.Print "Pagina "; Nr; " start "; p.Start; " eind "; p.End
        Nr = Nr + 1

        If Pos = 0 Then Exit Do

        StartPos = p.End + 1

    Loop

    Set LeesBoekPaginas = Paginas

End Function


'========================================================
' MAAK KATERN ARRAY
'========================================================

Private Sub MaakKaternArray( _
        ByVal EerstePagina As Long, _
        ByVal n As Long, _
        ByVal BoekPaginas As Collection)

    Dim p As Long
    Dim rij As Long
    Dim kol As Integer
    Dim Pagina As Long

    ReDim Katern(1 To 2 * n, 1 To 2)

 '   Debug.Print
     
   '  Debug.Print "KATERN ARRAY"

    For p = 1 To (PAGINAS_PER_STROOK * n)

        Pagina = EerstePagina + p - 1

        If Pagina <= BoekPaginas.Count Then

            KaternAdres _
                p, _
                n, _
                rij, _
                kol

   '         Debug.Print "Pagina "; Pagina; " -> "; rij; ","; kol

            Set Katern(rij, kol) = BoekPaginas(Pagina)

        End If

    Next p

End Sub


'========================================================
' KATERN ADRES
'========================================================

Private Sub KaternAdres( _
        ByVal p As Long, _
        ByVal n As Long, _
        ByRef rij As Long, _
        ByRef kol As Integer)

    Dim T As Long
    Dim s As Long

    T = PAGINAS_PER_STROOK * n

    If p Mod 2 = 1 Then

        If p <= 2 * n - 1 Then
            s = (p + 1) \ 2
            rij = 2 * s
            kol = 2
        Else
            s = (T - p + 1) \ 2
            rij = 2 * s - 1
            kol = 2
        End If

    Else

        If p <= 2 * n Then
            s = p \ 2
            rij = 2 * s - 1
            kol = 1
        Else
            s = (T - p + 2) \ 2
            rij = 2 * s
            kol = 1
        End If

    End If

End Sub


'========================================================
' DOELDOCUMENT UIT TEMPLATE
' iedere strook bestaatuit twee rijen. het aantal stroken is dus katernrijen /2
' het document heeft np stroken per pagina. er zijn dus nstroken\np paginas afgerond naar boven
'dat maakt het aaantal effectieve rijen TotaalRijen = ((TotaalRijen + np - 1) \ np) * np
'========================================================

Private Function MaakDoelDocument( _
        ByVal TemplatePad As String, _
        ByVal KaternRijen As Long, _
        ByRef np As Long) As Document

    Dim Doc As Document
    Dim Tabel As Table
    Dim Totaalrijen As Long

    
    Set Doc = Documents.Add(Template:=TemplatePad)

    If Doc.Tables.Count = 0 Then
        MsgBox "Template bevat geen tabel"
        Exit Function
    End If

    Set Tabel = Doc.Tables(1)
    np = Tabel.Rows.Count
    Totaalrijen = ((KaternRijen / 2 + np - 1) \ np) * np * 2
    
    Debug.Print "np template = "; np

    Do While Tabel.Rows.Count < Totaalrijen
        Tabel.Rows.Add
    Loop

    Debug.Print "doeltabel rijen = "; Tabel.Rows.Count

    Set MaakDoelDocument = Doc

End Function


'========================================================
' KOPIEER KATERN ARRAY NAAR DOELTABEL
'========================================================

Private Sub KopieerKaternNaarDoel( _
        ByVal DoelTabel As Table, _
        ByVal np As Long, _
        ByVal Totaalrijen As Long)

    Dim bronRij As Long
    Dim doelRij As Long
    Dim doelPagina As Long
    Dim rijOpPagina As Long
    Dim totaalPaginas As Long
    Dim nstroken As Long
    
    nstroken = Totaalrijen / 2
 
    ' Naar boven afronden in pure VBA:
    totaalPaginas = -Int(-nstroken / np) * 2
    

    
    Debug.Print
    Debug.Print "KOPIEER NAAR DOEL (Universele mapping voor np = " & np & ")"

    ' 1. EERST DE ONEVEN BRONRIJEN -> even doelpagina's
    doelPagina = 2
    rijOpPagina = 1
    
    For bronRij = 1 To Totaalrijen Step 2
        doelRij = ((doelPagina - 1) * np) + rijOpPagina
        If bronRij <= Totaalrijen And doelRij <= totaalPaginas * np Then
            KopieerKaternRij bronRij, doelRij, DoelTabel
            
            Dim berekendePagina As Long
            Dim berekendeRijOpPagina As Long
            
            berekendePagina = Int((doelRij - 1) / np) + 1
            berekendeRijOpPagina = ((doelRij - 1) Mod np) + 1
            
            Debug.Print "kopieer rij: "; bronRij; " NAAR PAGINA :"; berekendePagina; " RIJ: "; berekendeRijOpPagina
        End If
        doelPagina = doelPagina + 2
        If doelPagina > totaalPaginas Then
            doelPagina = 2
            rijOpPagina = rijOpPagina + 1
        End If
    Next bronRij
    
    ' 2. DAARNA DE EVEN BRONRIJEN -> oneven doelpagina's
    doelPagina = 1
    rijOpPagina = 1
    
    For bronRij = 2 To Totaalrijen Step 2
        doelRij = ((doelPagina - 1) * np) + rijOpPagina
        If bronRij <= Totaalrijen And doelRij <= totaalPaginas * np Then
            KopieerKaternRij bronRij, doelRij, DoelTabel
            
     
            berekendePagina = Int((doelRij - 1) / np) + 1
            berekendeRijOpPagina = ((doelRij - 1) Mod np) + 1
            
            Debug.Print "kopieer rij: "; bronRij; " NAAR PAGINA :"; berekendePagina; " RIJ: "; berekendeRijOpPagina
        End If
        doelPagina = doelPagina + 2
        If doelPagina > totaalPaginas Then
            doelPagina = 1
            rijOpPagina = rijOpPagina + 1
        End If
    Next bronRij

End Sub


'========================================================
' KOPIEER EEN KATERN RIJ
'========================================================

Private Sub KopieerKaternRij( _
        ByVal bronRij As Long, _
        ByVal doelRij As Long, _
        ByVal DoelTabel As Table)

    Dim c As Long

    For c = 1 To 2
        If Not Katern(bronRij, c) Is Nothing Then
            DoelTabel.Cell(doelRij, c).Range.FormattedText = Katern(bronRij, c).FormattedText
        End If
    Next c

    Debug.Print "Bron rij "; bronRij; " -> doel rij "; doelRij

End Sub


'========================================================
' TEMPLATE PAD (Automatisch zoeken)
'========================================================

Private Function TemplateNaam(ByVal BestandsNaam As String) As String

    Dim FSO As Object
    Dim ZoekPad As String
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If Not ActiveDocument Is Nothing Then
        If ActiveDocument.Path <> "" Then
            ZoekPad = ActiveDocument.Path & "\" & BestandsNaam
            If FSO.FileExists(ZoekPad) Then
                TemplateNaam = ZoekPad
                Exit Function
            End If
        End If
    End If
    
    On Error Resume Next
    If ThisDocument.Path <> "" Then
        ZoekPad = ThisDocument.Path & "\" & BestandsNaam
        If FSO.FileExists(ZoekPad) Then
            TemplateNaam = ZoekPad
            Exit Function
        End If
    End If
    On Error GoTo 0
    
    ZoekPad = Options.DefaultFilePath(wdUserTemplatesPath) & "\" & BestandsNaam
    If FSO.FileExists(ZoekPad) Then
        TemplateNaam = ZoekPad
        Exit Function
    End If
    
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Selecteer het sjabloon: " & BestandsNaam
    fd.Filters.Clear
    fd.Filters.Add "Word Sjablonen", "*.dotx; *.dotm; *.dot"
    
    If fd.Show = -1 Then
        TemplateNaam = fd.SelectedItems(1)
    Else
        MsgBox "Geen sjabloon geselecteerd.", vbCritical, "Fout"
        TemplateNaam = ""
    End If

End Function


'========================================================
' DOELMAP
'========================================================

Private Function MaakDoelMap(ByVal Doc As Document) As String
    Dim Naam As String
    Dim FSO As Object
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If Doc.Path = "" Then
        MsgBox "Sla het actieve document eerst op.", vbExclamation, "Niet opgeslagen"
        End
    End If
    
    Naam = Left(Doc.name, InStrRev(Doc.name, ".") - 1)
    MaakDoelMap = Doc.Path & "\" & Naam
    
    If Not FSO.FolderExists(MaakDoelMap) Then
        FSO.CreateFolder MaakDoelMap
    End If
End Function


'========================================================
' TEST KATERNADRES
'========================================================

Public Sub TestKaternAdres()
    Dim p As Long, r As Long
    Dim c As Integer

    Debug.Print "TEST KATERNADRES"
    For p = 1 To 32
        KaternAdres p, 8, r, c
        Debug.Print "Pagina "; p; " rij "; r; " kol "; c
    Next p
End Sub

