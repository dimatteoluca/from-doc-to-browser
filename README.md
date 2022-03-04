# FromDocToBrowser

Per evitare l'errore "ModuleNotFoundError: No module named 'textract.parsers.docx_parser'" dopo la conversione del codice in .exe tramite l'uso di programmi come "auto-py-to-exe"
è necessario aggiungere i parser "textract.parsers.docx_parser" e "textract.parsers.odt_parser" tramite la funzione --hidden-import che si trova tra le impostazioni avanzate.
