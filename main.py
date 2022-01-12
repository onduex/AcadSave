from pathlib import Path
import win32com.client
import time

filelist = []
i = 0

for path in Path('D:/$_INOVYN/Dossier 3-Vinilis-PVC I').glob('**/*.dwg'): # THIS IS OK 2602
# for path in Path('D:/$_INOVYN/Dossier 1-Hispavic Ibérica-UE').glob('**/*.dwg'): THIS IS OK 2190
# for path in Path('D:/$_INOVYN/Dossier 0-Trabajos externos-Especificaciones').glob('**/*.dwg'): THIS IS OK 350
    filelist.append(path)


for record in filelist:
    acad = win32com.client.dynamic.Dispatch("AutoCAD.Application")
    time.sleep(1)
    acad.Visible=False
    
    # Open a new document and set it as the active document
    acad.Documents.Open(record)
    time.sleep(1)
    # Set the active document before trying to use it
    doc = acad.ActiveDocument
    time.sleep(1)
    ### Adjust dwg ###
    doc.Save()
    doc.Close()
    i += 1
    print(i, len(filelist), record)
