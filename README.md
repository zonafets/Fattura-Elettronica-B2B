# Fattura Elettronica B2B

** Questo codice è stato verificato e ha prodotto fatture accettate dal SID con IVA e IVA esente, dettagli ordinari e dettagli scontati. **

### Scopo
Essendo il BASIC un linguaggio molto semplice, il seguente codice fornisce un modello propedeutico.

### Classe "efattura"

I dati del software d'origine sono stati precedentemente normalizzati tramite query.

La funzione impiega la libreria Microsoft MSXML6 ma attraverso le funzioni "into" e "add", è possibile generare da sè il documento, evitendo tale dipendenza.

#### Dichiarazioni

```vba
' da tabella "efatture_costanti"
Public Enum TipoDocumentoType
    Fattura = 1
    NotaDiCredito = 4
End Enum

Private Const errNotMarked = 1
Private Const errProgressive = 2
Private Const errXMLStruct = 3
Private Const errXMLAdd = 4
Private Const errXMLParse = 5
Private Const errNoTarget = 6

' se da errore in questa riga (manca DOMDocument60):
' 1. stoppare l'esecuzione di ogni programma
' 2. aprire un modulo vbasic
' 3. Strumenti->Riferimenti cercare e aggiungere "Microsoft XML, v6.0"

Private doc As DOMDocument60
Private root As IXMLDOMElement
Private tabella
```

#### Funzioni di comodo
```vba
Private Function cfg(ID As String) As String
    cfg = econfig(ID)
End Function
```

```vba
Private Function progressivo() As Integer
    progressivo = CInt(cfg("progressivo"))
End Function
```

```vba
Private Sub avanzaProgressivo()
    Dim db
    Set db = Application.CurrentDb
    db.Execute ("update efattura_variabili set valore=valore+1 where id='progressivo'")
    If db.recordsaffected = 0 Then Err.Raise errProgressive, "efattura", "Progressivo non avanzato"
End Sub
```

```vba
Private Sub registraDataProgressivoXML(Tipo As TipoDocumentoType, ID As Long, progressivo As String)
    Dim db
    Set db = Application.CurrentDb
    
    Select Case Tipo
        
        Case Fattura:
        db.Execute ("update tabella_fatturazione set dataXML=now(),progressivoXML='" & progressivo & "' where id_fattura=" & ID)
        
        Case NotaDiCredito:
        db.Execute ("update tabella_note_accredito set dataXML=now(),progressivoXML='" & progressivo & "' where [id_nota-accredito]=" & ID)

    End Select
    
    If db.recordsaffected = 0 Then Err.Raise errNotMarked, "efattura", "Non è stato possibile marcare la data(XML) fattura"
End Sub
```

```vba
Private Function FileExists(ByVal path_ As String) As Boolean
    FileExists = (Len(Dir(path_)) > 0)
End Function
```

```vba
' ==========================================================
' facility per gestione nodi xml
' ==========================================================

Private Sub into(name As String, Optional tabellaDiRiferimento)
    Dim child As IXMLDOMElement
    Set child = doc.createElement(name)
    root.appendChild child
    Set root = child
    If Not IsMissing(tabellaDiRiferimento) Then Set tabella = tabellaDiRiferimento
    'ReDim Preserve xmlPath(xmlPathIdx)
    'xmlPath(xmlPathIdx) = name
    'xmlPathIdx = xmlPathIdx + 1
End Sub

Private Sub out(Optional currentNodeName)
    If Not IsMissing(currentNodeName) Then
        If root.nodeName <> currentNodeName Then
            Err.Raise errXMLStruct, "efattura", "Navigazione errata nell'XML:" + root.nodeName + "<>" + currentNodeName
        End If
    End If
    Set root = root.parentNode
    'ReDim Preserve xmlPath(xmlPathIdx - 1)
    'xmlPathIdx = xmlPathIdx - 1
End Sub

Private Sub add(name As String, Optional value)
    Dim child As IXMLDOMElement
    Dim err_number As Long
    
    If IsMissing(value) And Not tabella Is Nothing Then
        On Error Resume Next
        value = tabella.Fields(name)
        err_number = Err.Number
        On Error GoTo 0
        If err_number <> 0 Then Err.Raise errXMLAdd, "efattura", "Campo '" + name + "' non trovato nella tabella '" + tabella.name + "'"
    End If
    
    If Nz(value, "") = "" Then Exit Sub
    
    Set child = doc.createElement(name)
    root.appendChild child
    child.Text = value
End Sub
```

```vba
' ==========================================================
' utilità per stampa indentata dell'xml
' ==========================================================

' questo stampa utf-16 anziché 8
Private Function PrettyXML(ByVal Source, Optional ByVal EmitXMLDeclaration As Boolean) As String
    Dim Writer As MXXMLWriter60, Reader As SAXXMLReader60
    Set Writer = New MXXMLWriter60
    Writer.indent = True
    Writer.omitXMLDeclaration = Not EmitXMLDeclaration
    Set Reader = New SAXXMLReader60
    Set Reader.contentHandler = Writer
    Reader.parse Source
    PrettyXML = Writer.Output
End Function
```

```vba
Private Function dati(tabella As String, Optional riferimenti As Variant, Optional chiavi As Variant)
    Dim sql As String
    Dim valore As String
    Dim chiave As Variant
    Dim riferimento As String
    Dim ariferimenti As Variant
    Dim achiavi As Variant
    Dim where As String
    Dim i As Integer
    If IsMissing(riferimenti) Then
        sql = "select * from [" + tabella + "]"
    Else
        If TypeName(riferimenti) <> "Variant()" Then
            ariferimenti = Array(riferimenti)
        Else
            ariferimenti = riferimenti
        End If
        If TypeName(chiavi) <> "Variant()" Then
            achiavi = Array(chiavi)
        Else
            achiavi = chiavi
        End If
        
        For i = LBound(ariferimenti) To UBound(ariferimenti)
            chiave = achiavi(i)
            riferimento = ariferimenti(i)
            If TypeName(chiave) = "Field" Then chiave = chiavi(i).valore
            If TypeName(chiave) = "string" Then valore = """" + chiave + """" Else valore = CStr(chiave)
            where = where + IIf(i > 0, " and ", "") + "([" + riferimento + "]=" + valore + ")"
        Next i
        sql = "select * from [" + tabella + "] where " + where
        If tabella <> "efattura_cliente" Then
            sql = sql + " order by Numero"
        End If
    End If
    Set dati = Application.CurrentDb.OpenRecordset(sql)
End Function
```

#### Funzione principale
```vba

' ==========================================================
' genera files xml per fatturazione elettronica
' ==========================================================

Private Function FatturaElettronica(ID As Long, Tipo As TipoDocumentoType) As String

    ' ===========================================
    ' documentazione ufficiale
    ' ===========================================
    '
    ' http://www.fatturapa.gov.it/export/fatturazione/sdi/fatturapa/v1.2.1/Rappresentazione_tabellare_del_tracciato_FatturaPA_versione_1.2.1.xls
    ' http://www.fatturapa.gov.it/export/fatturazione/sdi/Suggerimenti_Compilazione_FatturaPA_1.5.pdf
    '
    ' fogli di stile per stampe
    ' ordinaria: http://www.fatturapa.gov.it/export/fatturazione/sdi/fatturapa/v1.2.1/fatturaordinaria_v1.2.1.xsl
    ' verso PA:  http://www.fatturapa.gov.it/export/fatturazione/sdi/fatturapa/v1.2.1/fatturaPA_v1.2.1.xsl
    
    ' legge le testate fino al pregresso di 240gg
    Dim cliente
    Dim testata
    Dim aliquote
    Dim dettaglio
    Dim node As IXMLDOMElement
    Dim path_destinazione As String
    
    Dim num As Long
    Dim TipoDocumento As String
    
    Dim tabella_testata As String
    Dim tabella_dettaglio As String
    Dim tabella_iva As String
    Dim testata_chiavi As Variant
    Dim testata_valori As Variant
    
    path_destinazione = cfg("path_deposito_efattura")
    
    TipoDocumento = "TD" & Right("00" & Tipo, 2)
    
    If (Tipo = Fattura) Then
        tabella_testata = "efattura_fattura"
        tabella_dettaglio = "efattura_fattura_dettaglio"
        tabella_iva = "efattura_fattura_iva"
    Else
        tabella_testata = "efattura_nota_di_credito"
        tabella_dettaglio = "efattura_nota_di_credito_dettaglio"
        tabella_iva = "efattura_nota_di_credito_iva"
    End If
    num = ID
        
    Set doc = New DOMDocument60
    ' doc.async = False
    ' doc.SetProperty "ProhibitDTD", False
    ' doc.validateOnParse = False
    ' usare il solo <?xml version="1.0" encoding="UTF-8"?> da errore perché richiede la root
    ' doc.loadXML "<?xml version=""1.0"" encoding=""UTF-8""?>"
    
    doc.loadXML cfg("root-efattura-1.2.1")
    doc.SetProperty _
        "SelectionNamespaces", _
        "xmlns:p='http://ivaservizi.agenziaentrate.gov.it/docs/xsd/fatture/v1.2'"
    
    ' ma per questioni di namespace, e:FatturaElettronica non è selezionabile con SelectSingleNode("p:...")
    If doc.parseError.reason <> "" Then Err.Raise errXMLParse, "efattura", doc.parseError.reason
        
    Set root = doc.selectSingleNode("p:FatturaElettronica")
    
    Set testata = dati(tabella_testata, "Numero", num)
    
    Set cliente = dati("efattura_cliente", "ID", testata!ID_cliente)

    ' questa sezione può essere espressa una sola volta per più fatture/note per lo stesso cliente

    into "FatturaElettronicaHeader", cliente
    
        into "DatiTrasmissione"
        
            into "IdTrasmittente"
                add "IdPaese", "IT"
                add "IdCodice", cfg("cf") ' cfg("piva")
                out
            
            add "ProgressivoInvio", CStr(progressivo())
            add "FormatoTrasmissione", "FPR12"
            
            If (Nz(tabella!CodiceDestinatario, "") = "" And Nz(tabella!PECDestinatario, "") = "") Then
                Err.Raise errNoTarget, "efattura", "CodiceDestinatario e PEC mancenti per cliente '" & tabella!Denominazione & "'"
            End If
            
            add "CodiceDestinatario"                    ' vedere commento in modulo "efatture_utilità"
            add "PECDestinatario"                       ' ???? Come gestire la mancanza?
        
            out
        
        into "CedentePrestatore"
        
            into "DatiAnagrafici"
                into "IdFiscaleIVA"
                    add "IdPaese", "IT"
                    add "IdCodice", cfg("piva")
                    out
                into "Anagrafica"
                    add "Denominazione", cfg("denominazione")
                    out
                add "RegimeFiscale"
                out
                
            into "Sede"
                add "Indirizzo", cfg("indirizzo")
                add "CAP", cfg("cap")
                add "Comune", cfg("comune")
                add "Provincia", cfg("provincia")
                add "Nazione", cfg("nazione")
                out
            out
        
        into "CessionarioCommittente"
        
            into "DatiAnagrafici"
                add "CodiceFiscale"
                into "Anagrafica"
                    add "Denominazione"
                    out
                out
                
            into "Sede"
                add "Indirizzo"
                add "CAP"
                add "Comune"
                add "Provincia"
                add "Nazione"
                out
            out
            
        out "FatturaElettronicaHeader"
        
    cliente.Close
    Set cliente = Nothing
    
    ' questa sezione può essere ripetuta se le fatture/note appartengono allo stesso cliente
        
    into "FatturaElettronicaBody", testata
    
        into "DatiGenerali"
        
            into "DatiGeneraliDocumento"
            
                add "TipoDocumento", TipoDocumento
                add "Divisa"
                add "Data"
                add "Numero"
                ' la causale, se c'è, va spezzata in blocchi da 200 caratteri
                out
                
            ' into "DatiOrdineAcquisto"  ' non obbligatorio
            ' into "DatiContratto"       ' non obbligatorio
            ' into "DatiTrasporto"       ' non obbligatorio
            If Tipo = NotaDiCredito Then
                into "DatiFattureCollegate"
                    add "IdDocumento", testata!IDFatturaCollegata
                    add "Data", testata!DataFatturaCollegata
                    out
            End If
            
            out
               
        into "DatiBeniServizi"
        
            Set dettaglio = dati(tabella_dettaglio, "Numero", testata!Numero)
                
            ' la numerazione originale segue il servizio anziché l'intera fattura
            ' quindi viene rigenerata manualmente, poiché via sql sarebbe dispendioso
            Dim NumeroLinea As Integer
            NumeroLinea = 0
            
            While Not dettaglio.EOF
                
                NumeroLinea = NumeroLinea + 1
                
                into "DettaglioLinee", dettaglio
                    add "NumeroLinea", NumeroLinea
                    ' la descrizione è limitata a 100 caratteri
                    add "Descrizione", Left(dettaglio!Descrizione, 100)
                    add "Quantita"
                    add "UnitaMisura"
                    add "PrezzoUnitario"
                    
                    If dettaglio!Percentuale <> "" Or dettaglio!Importo <> "" Then
                        into "ScontoMaggiorazione"
                            add "Tipo"
                            If dettaglio!Percentuale <> "" Then add "Percentuale"
                            ' se specificati assieme, "Importo" prevale
                            If dettaglio!Importo <> "" Then add "Importo"
                        out
                    End If
                    
                    add "PrezzoTotale"
                    add "AliquotaIVA"
                    
                    If dettaglio!Natura <> "" Then
                        add "Natura"
                    End If
                    
                    out
                                
                dettaglio.MoveNext
            Wend ' dettaglio
        
            Set aliquote = dati(tabella_iva, "Numero", testata!Numero)
            
            While Not aliquote.EOF
                                
                into "DatiRiepilogo", aliquote
                    add "AliquotaIVA"
                    
                    If aliquote!Natura <> "" Then
                        add "Natura"
                    End If

                    add "ImponibileImporto"
                    add "Imposta"
                    add "EsigibilitaIVA"
                    out

                aliquote.MoveNext
            Wend ' aliquote
            
            out "DatiBeniServizi"
        
        into "DatiPagamento", testata
            add "CondizioniPagamento"
            into "DettaglioPagamento"
                add "ModalitaPagamento"
                add "DataScadenzaPagamento"
                add "ImportoPagamento"
                out
            out
            
        registraDataProgressivoXML Tipo, testata!Numero, CStr(progressivo())

        out "FatturaElettronicaBody"
            
    ' Debug.Print PrettyXML(doc, True)
    
    Dim nome_xml As String
    Dim file As String
    nome_xml = "IT" + cfg("piva") + "_" + Format(progressivo, "00000") & ".xml"
    file = Replace(Replace(path_destinazione, ".", CurDir) + "\", "\\", "\") + nome_xml
    If (FileExists(file)) Then Kill (file)
    
    ' anche se debug.print doc.xml stampa solo <?xml version="1.0"?>
    ' ciene salvato correttamente con ... encoding="UTF-8"?>
    doc.Save file
    
    Debug.Print "Generata fattura elettronica in: " + file
    
    FatturaElettronica = file
    
    avanzaProgressivo
    
    Set tabella = Nothing

End Function ' FatturaElettronica

```

#### Funzione interfaccia esterna
```vba
Public Function Genera(ID As Long, Tipo As TipoDocumentoType) As String

    Application.DBEngine.BeginTrans
    On Error GoTo errors
    Genera = FatturaElettronica(ID, Tipo)
    Application.DBEngine.CommitTrans
    Exit Function
    
errors:
    Application.DBEngine.Rollback
    MsgBox Err.Description, vbError, "Errore nell'esportazione"
    On Error GoTo 0

End Function
```

#### Esempio d'uso
```vba
Public Function efatturaGeneraFattura(ID_Fattura As Long)

    Dim fatturaxml As New efattura
    Dim file As String
    file = fatturaxml.Genera(ID_Fattura, Fattura)
    If file <> "" Then MsgBox "Generata fattura " & file
    Set fatturaxml = Nothing
    
End Function

Public Function efatturaGeneraNota(ID_Nota_Di_Credito As Long)

    Dim fatturaxml As New efattura
    Dim file As String
    file = fatturaxml.Genera(ID_Nota_Di_Credito, NotaDiCredito)
    If file <> "" Then MsgBox "Generata nota di accredito " & file
    Set fatturaxml = Nothing

End Function
```
