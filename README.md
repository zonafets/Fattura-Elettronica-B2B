# Fattura Elettronica B2B

** Questo codice è stato verificato e ha prodotto fatture accettate dal SID con IVA e IVA esente, dettagli ordinari e dettagli scontati, bollo. **

### Scopo
Essendo il BASIC un linguaggio molto semplice, il seguente codice fornisce un modello propedeutico.

### Classe "efattura"

I dati del software d'origine sono stati precedentemente normalizzati tramite query.

La funzione impiega la libreria Microsoft MSXML6 ma attraverso le funzioni "into" e "add", è possibile generare da sè il documento, evitendo tale dipendenza.


#### Funzioni di comodo
```vba
Private Function cfg(ID As String) As String

Private Function progressivo() As Integer

Private Sub avanzaProgressivo()

Private Sub registraDataProgressivoXML(Tipo As 

Private Function FileExists(ByVal path_ As String) As Boolean
```

```vba
' ==========================================================
' facility per gestione nodi xml
' ==========================================================

Private Sub into(name As String, Optional tabellaDiRiferimento)

Private Sub out(Optional currentNodeName)

Private Sub add(name As String, Optional value)
```

```vba
' ==========================================================
' utilità per stampa indentata dell'xml
' ==========================================================

' n.b. questo stampa utf-16 anziché 8
Private Function PrettyXML(ByVal Source, Optional ByVal EmitXMLDeclaration As Boolean) As String
```

```vba
' imposta la tabella da cui acquisire i dati
Private Function dati(tabella As String, Optional riferimenti As Variant, Optional chiavi As Variant)
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
                    add "IdCodice", cfg("cf") ' 190211\s.zaglio ex piva
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
                If (Nz(testata!ImportoBollo, "") <> "") Then
                    into "DatiBollo"
                        add "BolloVirtuale", "SI"
                        add "ImportoBollo"
                        out
                End If

                add "Causale", testata!causale
                add "Causale", testata!causale1
                add "Causale", testata!causale2
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
                    
                    add "Natura" ' valorizzata solo se AIVA = 0
                    
                    out
                                
                dettaglio.MoveNext
            Wend ' dettaglio
        
            Set aliquote = dati(tabella_iva, "Numero", testata!Numero)
            
            While Not aliquote.EOF
                                
                into "DatiRiepilogo", aliquote
                    add "AliquotaIVA"
                    
                    add "Natura" ' valorizzata solo se AIVA = 0
                    add "RiferimentoNormativo" ' valorizzato solo se AIVA = 0
                    
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
```

#### Esempio d'uso
```vba
Public Function efatturaGeneraFattura(ID_Fattura As Long)

Public Function efatturaGeneraNota(ID_Nota_Di_Credito As Long)
```
