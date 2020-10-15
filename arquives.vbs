OPTION EXPLICIT
DIM fileSystem, folder, file, path, objXMLDoc, Root, proPred, vCarga, CargaValue

path = "docs\"
SET fileSystem = CreateObject("Scripting.FileSystemObject")
SET folder = fileSystem.GetFolder(path)

FOR EACH file in folder.Files    

  IF checkArquive("docs\"& file.name) THEN
    IF DateDiff("M", file.DateLastModified, Now) < 1 THEN
      IF FileExists("docs\importado\"& file.name) THEN 
      ELSE
        filesystem.CopyFile "docs\" & file.name , "docs\averbar\"
        filesystem.CopyFile "docs\" & file.name , "docs\importado\"
      END IF
    END IF
  END IF

NEXT


FUNCTION FileExists(FilePath)
  IF fileSystem.FileExists(FilePath) THEN
    FileExists = CBool(1)
  ELSE
    FileExists = CBool(0)
  END IF
END FUNCTION 


FUNCTION checkArquive(arquive)
  SET objXMLDoc = CreateObject("Microsoft.XMLDOM")
  objXMLDoc.async = False 
  objXMLDoc.load(arquive)

  SET Root = objXMLDoc.documentElement 
  SET vCarga = Root.getElementsByTagName("vCarga")
  SET proPred = Root.getElementsByTagName("proPred")
  SET CargaValue = vCarga.CDec()
  
  IF proPred(0).text = "INSUMO PARA FUNDICAO" AND vCarga(0).text < 1000000 THEN
      checkArquive = CBool(1)
    ELSEIF proPred(0).text <> "INSUMO PARA FUNDICAO" AND vCarga(0).text < 300000 THEN
        checkArquive = CBool(1)
      ELSE
        checkArquive = CBool(0)
  END IF

END FUNCTION 