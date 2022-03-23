from os import listdir
from os.path import isfile, isdir
import re

print("Hello World")
files = []
ignore = ['ThisDocument', 'VBAReset', 'vw_base_signal_c', 'vw_err']

def GetComps(path):
    for f in listdir(path):
        if re.match("\.", f):
            continue
        elif isfile(path + "/" + f) and re.search("\.bas$|\.cls$|\.frm$", f):
            if not re.match("|".join(ignore), f):
                files.append(path + "\\" + f)
        elif isdir(f):
            GetComps(path + "\\" + f)

GetComps("D:\VW")

fout = open('VBAReset.bas', 'w')
fout.write("Attribute VB_Name = \"VBAReset\"\n\n")
fout.write("' Generated by Python Script! Edit iif you know what you are doing\n")
fout.write("' Use after enabling Macro Settings -> Allow Programmatic Access to VBProject\n\n")
fout.write("Public Sub VBA_Reset()\n")
fout.write("  On Error Resume Next\n")
fout.write("  Dim MyComponents as Collection\n")
fout.write("  Set MyComponents = New Collection\n")
for f in files:
    fout.write("  MyComponents.Add \"" + f + "\"\n")
fout.write("  Do While ThisDocument.VBProject.VBComponents.Count > 2\n")
fout.write("    For Each vbComp in ThisDocument.VBProject.VBComponents\n")
fout.write("      If Left$(vbComp.Name, 3) = \"vw_\" Then _\n")
fout.write("        ThisDocument.VBProject.VBComponents.Remove vbComp\n")
fout.write("    Next\n")
fout.write("  Loop\n")
fout.write("  For Each vbComp in MyComponents\n")
fout.write("    Application.VBE.ActiveVBProject.VBComponents.Import vbComp\n")
fout.write("  Next\n")
fout.write("End Sub")
fout.close()