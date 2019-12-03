OnErrorGoto(?CatchError)

Enumeration ;- Fenster IDs
  #Win_Main
  #Win_AddFile
  #Win_Log
EndEnumeration

Enumeration ;- Gadget IDs
  #LstVw_Filelist
  #Txt_Filelist
  #F3D_Nummerierung
  #Txt_Name
  #Stg_Name
  #Stg_NrStart
  #Stg_NrFactor
  #Txt_NrStart
  #Txt_NrFactor
  #F3D_ReplaceText
  #Stg_Replace_1
  #Stg_With_1
  #Stg_Replace_2
  #Stg_With_2
  #Txt_Replace
  #Txt_With
  #F3D_Typ
  #Rad_Typ_Copy
  #Rad_Typ_RenameOnly
  #F3D_EditDate
  #ChkBox_CreateDate
  #ChkBox_LastAcces
  #ChkBox_LastEdit
  #Cmb_LastAcces
  #Cmb_CreateDate
  #Cmb_LastEdit
  #F3D_EditAttribute
  #ChkBox_Hide
  #ChkBox_ReadOnly
  #ChkBox_System
  #ChkBox_Normal
  #F3D_Filearea
  #Btn_Help
  #Btn_Info
  #Btn_Copyfolder
  #Btn_AddFiles
  #Btn_RemoveFiles
  #Btn_ClearFiles
  #Btn_MusterInfo
  #ChkBox_ShowPath
  #Spn_ZeroAmount
  #Txt_ZeroAmount
  #ChkBox_RemoveNummeric
  #ChkBox_RemoveSpaces
  #ChkBox_RemoveMisc
  #Stg_RemoveMisc_6
  #Btn_Start
  #Stg_RemoveMisc_5
  #Stg_RemoveMisc_4
  #Stg_RemoveMisc_3
  #Stg_RemoveMisc_2
  #Stg_RemoveMisc_1
  #ChkBox_IgnoreReplace_1
  #ChkBox_IgnoreReplace_2
  #ChkBox_OverwriteMsg
  
  #ExpLst_SearchFile
  #Txt_SearchIn
  #Cmb_SearIn
  #Btn_AddFile
  #Btn_AddFileCancel
  #Cmb_FileTyp
  #Txt_FileTyp
  #Btn_ExportFolder
  #Btn_Ansicht

  #Stg_Log
  #Btn_SaveLog
  #Btn_CloseLog
  #Btn_KopieLog
  #Btn_StopRenameEvent
  #PrgBar_RenameEvent
  #Txt_RenameEvent
  #Btn_OpenFolder
EndEnumeration

Enumeration ;- Menu IDs
  #Ansicht
  #Ansicht_Symbole
  #Ansicht_List
  #Ansicht_Details
  
  #EnterFileTyp
EndEnumeration

Procedure WCB(wnd, Message, wParam, lParam)
 Result = #PB_ProcessPureBasicEvents
 Select Message
  Case #WM_GETMINMAXINFO
   GetWindowRect_(wnd,r.RECT)
   *pMinMax.MINMAXINFO = lParam
   If wnd = WindowID(#Win_AddFile)
    *pMinMax\ptMinTrackSize\x= 458 ;Breite + 8
    *pMinMax\ptMinTrackSize\y= 290 ;Höhe + 32
   ElseIf wnd = WindowID(#Win_Log)
    *pMinMax\ptMinTrackSize\x= 486 ;Breite + 8
    *pMinMax\ptMinTrackSize\y= 275 ;Höhe + 27
   EndIf
   *pMinMax\ptMaxTrackSize\x=GetSystemMetrics_(#SM_CXSCREEN)
   *pMinMax\ptMaxTrackSize\y=GetSystemMetrics_(#SM_CYSCREEN)
   Result = 0
 EndSelect
 ProcedureReturn Result
EndProcedure

;Konstanten
#Programmname   = "Renamer"
#Programmvers   = "1.00"
#nL             = Chr(13)+Chr(10)

;Variablen
Global Entr.s, Name.s, Exts.s, StNr.l, ErNr.l, AkNr.l, Prgs.l, Prga.l, Amon.l, ErCo.l, Resu.l
Global thread.l
Global RenameEventT.l  = #False
Global SaveFolder.s    = ExePath() + "Export\"
Global SearchFolder.s  = ExePath()
Global NewList FileList.s()
Global NewList DelPos.l()

If OpenWindow(#Win_Main, 0, 0, 540, 390, #Programmname,  #PB_Window_SystemMenu|#PB_Window_MinimizeGadget|#PB_Window_TitleBar|#PB_Window_ScreenCentered|#PB_Window_Invisible)
 If CreateGadgetList(WindowID(#Win_Main))
  ListViewGadget(#LstVw_Filelist, 5, 25, 295, 265, #PB_ListView_ClickSelect)
  TextGadget(#Txt_Filelist, 5, 10, 180, 15, "0 Dateien")
  Frame3DGadget(#F3D_Nummerierung, 310, 40, 230, 55, "Nummerierung")
  TextGadget(#Txt_Name, 310, 5, 200, 15, "Name (Muster):")
  StringGadget(#Stg_Name, 310, 20, 205, 20, "Datei $Z")
  StringGadget(#Stg_NrStart, 320, 70, 60, 20, "", #PB_String_Numeric)
  StringGadget(#Stg_NrFactor, 390, 70, 60, 20, "", #PB_String_Numeric)
  TextGadget(#Txt_NrStart, 320, 55, 60, 15, "Start")
  TextGadget(#Txt_NrFactor, 390, 55, 60, 15, "Erhöhung")
  Frame3DGadget(#F3D_ReplaceText, 310, 95, 230, 140, "Dateiname bearbeiten")
  StringGadget(#Stg_Replace_1, 320, 125, 80, 20, "")
  StringGadget(#Stg_With_1, 400, 125, 80, 20, "")
  StringGadget(#Stg_Replace_2, 320, 145, 80, 20, "")
  StringGadget(#Stg_With_2, 400, 145, 80, 20, "")
  TextGadget(#Txt_Replace, 320, 110, 80, 15, "Ersetzen")
  TextGadget(#Txt_With, 400, 110, 80, 15, "Mit")
  Frame3DGadget(#F3D_Typ, 0, 320, 240, 70, "Methode")
  OptionGadget(#Rad_Typ_Copy, 10, 340, 170, 15, "Dateien im Zielordner kopieren")
  OptionGadget(#Rad_Typ_RenameOnly, 10, 370, 170, 15, "Dateien nur umbenennen")
  Frame3DGadget(#F3D_EditDate, 310, 235, 230, 85, "Datum ändern")
  CheckBoxGadget(#ChkBox_CreateDate, 320, 253, 110, 15, "Erstelldatum")
  CheckBoxGadget(#ChkBox_LastAcces, 320, 274, 110, 15, "Letzter Zugriff")
  CheckBoxGadget(#ChkBox_LastEdit, 320, 297, 110, 15, "Letzte veränderung")
  DateGadget(#Cmb_LastAcces, 435, 272, 95, 20)
  DateGadget(#Cmb_CreateDate, 435, 250, 95, 20)
  DateGadget(#Cmb_LastEdit, 435, 295, 95, 20)
  Frame3DGadget(#F3D_EditAttribute, 310, 320, 230, 70, "Attribute ändern")
  CheckBoxGadget(#ChkBox_Hide, 320, 355, 105, 15, "Versteckt")
  CheckBoxGadget(#ChkBox_ReadOnly, 425, 340, 105, 15, "Schreibgeschützt")
  CheckBoxGadget(#ChkBox_System, 425, 355, 105, 15, "Systemdatei")
  CheckBoxGadget(#ChkBox_Normal, 320, 340, 105, 15, "Normal")
  Frame3DGadget(#F3D_Filearea, 0, 5, 305, 315, "", #PB_Frame3D_Double)
  ButtonGadget(#Btn_Help, 250, 330, 50, 25, "Hilfe")
  ButtonGadget(#Btn_Info, 250, 360, 50, 25, "Info")
  ButtonGadget(#Btn_Copyfolder, 195, 345, 30, 15, "...")
  ButtonGadget(#Btn_AddFiles, 10, 295, 68, 20, "Hinzufügen")
  ButtonGadget(#Btn_RemoveFiles, 80, 295, 68, 20, "Entfernen")
  ButtonGadget(#Btn_ClearFiles, 150, 295, 68, 20, "Leeren")
  ButtonGadget(#Btn_MusterInfo, 518, 20, 20, 20, "?")
  CheckBoxGadget(#ChkBox_ShowPath, 190, 10, 110, 15, "Vollständiger Path")
  SpinGadget(#Spn_ZeroAmount, 460, 70, 50, 20, 1, 20, #PB_Spin_Numeric|#PB_Spin_ReadOnly)
  TextGadget(#Txt_ZeroAmount, 460, 55, 50, 15, "Stellen")
  CheckBoxGadget(#ChkBox_RemoveNummeric, 430, 195, 100, 15, "Zahlen")
  CheckBoxGadget(#ChkBox_RemoveSpaces, 430, 210, 100, 15, "Leerzeichen")
  CheckBoxGadget(#ChkBox_RemoveMisc, 320, 175, 65, 15, "Entfernen")
  StringGadget(#Stg_RemoveMisc_6, 390, 210, 35, 20, "")
  ButtonGadget(#Btn_Start, 225, 295, 68, 20, "Start")
  StringGadget(#Stg_RemoveMisc_5, 355, 210, 35, 20, "")
  StringGadget(#Stg_RemoveMisc_4, 320, 210, 35, 20, "")
  StringGadget(#Stg_RemoveMisc_3, 390, 190, 35, 20, "")
  StringGadget(#Stg_RemoveMisc_2, 355, 190, 35, 20, "")
  StringGadget(#Stg_RemoveMisc_1, 320, 190, 35, 20, "")
  CheckBoxGadget(#ChkBox_IgnoreReplace_1, 482, 127, 48, 15, "AaZz")
  CheckBoxGadget(#ChkBox_IgnoreReplace_2, 482, 147, 48, 15, "AaZz")
  CheckBoxGadget(#ChkBox_OverwriteMsg, 15, 355, 170, 15, "Ohne Nachfrage überschreiben")
 EndIf
EndIf
SetGadgetState(#Rad_Typ_Copy, #True)
SetGadgetState(#Spn_ZeroAmount, 3)
SetGadgetText(#Stg_NrStart, "1")
SetGadgetText(#Stg_NrFactor, "1")
SendMessage_(GadgetID(#Stg_NrStart), #EM_SETLIMITTEXT, 8, 0)
SendMessage_(GadgetID(#Stg_NrFactor), #EM_SETLIMITTEXT, 8, 0)

If OpenWindow(#Win_AddFile, 0, 0, 450, 262, "Dateien hinzufügen", #PB_Window_SizeGadget|#PB_Window_WindowCentered|#PB_Window_SystemMenu|#PB_Window_Invisible, WindowID(#Win_Main))
 SetWindowCallback(@WCB(),#Win_AddFile)
 If CreateGadgetList(WindowID(#Win_AddFile))
  ExplorerListGadget(#ExpLst_SearchFile, 5, 30, 440, 200, SearchFolder, #PB_Explorer_MultiSelect|#PB_Explorer_NoDriveRequester|#PB_Explorer_AlwaysShowSelection)
  TextGadget(#Txt_SearchIn, 5, 8, 55, 20, "Suchen in:")
  ExplorerComboGadget(#Cmb_SearIn, 60, 5, 300, 250, SearchFolder)
  ButtonGadget(#Btn_AddFile, 300, 235, 70, 22, "Hinzufügen")
  ButtonGadget(#Btn_AddFileCancel, 375, 235, 70, 22, "Abbrechen")
  StringGadget(#Cmb_FileTyp, 55, 230, 240, 20, "*.*")
  TextGadget(#Txt_FileTyp, 5, 238, 45, 15, "Dateityp:")
  ExplorerComboGadget(#Cmb_SearIn, 60, 5, 280, 250, SearchFolder)
  ButtonGadget(#Btn_ExportFolder, 395, 5, 50, 20, "Export")
  ButtonGadget(#Btn_Ansicht, 343, 5, 50, 20, "Ansicht")
 EndIf
  If CreatePopupMenu(#Ansicht)
   MenuItem(#Ansicht_Symbole, "Symbole")
   MenuItem(#Ansicht_List,"Liste")
   MenuItem(#Ansicht_Details, "Details")
  EndIf
EndIf
ChangeListIconGadgetDisplay(#ExpLst_SearchFile, #PB_ListIcon_List)
AddGadgetItem(#Cmb_FileTyp, -1, "Alle Dateien (*.*)")
AddGadgetItem(#Cmb_FileTyp, -1, "Grafiken (*.bmp, *.png, *.jpg, *.gif)")
SetGadgetState(#Cmb_FileTyp, 0)

If OpenWindow(#Win_Log, 0, 0, 478, 248, "Protokoll",  #PB_Window_SizeGadget|#PB_Window_WindowCentered|#PB_Window_Invisible, WindowID(#Win_Main))
 SetWindowCallback(@WCB(),#Win_Log)
 If CreateGadgetList(WindowID(#Win_Log))
  EditorGadget(#Stg_Log, 0, 0, 370, 220, #ES_MULTILINE|#ES_AUTOVSCROLL|#ES_NOHIDESEL|#ES_READONLY)
  ButtonGadget(#Btn_SaveLog, 375, 40, 100, 20, "Speichern")
  ButtonGadget(#Btn_KopieLog, 375, 15, 100, 20, "Zwischenablage")
  ButtonGadget(#Btn_OpenFolder, 375, 65, 100, 20, "Ordner öffnen")
  ButtonGadget(#Btn_StopRenameEvent, 375, 200, 100, 20, "Stop")
  ButtonGadget(#Btn_CloseLog, 375, 225, 100, 20, "Schliessen")
  ProgressBarGadget(#PrgBar_RenameEvent, 0, 235, 370, 10, 0, 100)
  TextGadget(#Txt_RenameEvent, 0, 220, 370, 12, "Datei 0/0")
 EndIf
EndIf
ShowScrollBar_(GadgetID(#Stg_Log), #SB_BOTH, #True)
SetGadgetColor(#Stg_Log, #PB_Gadget_BackColor, RGB(255,255,255))

If FileSize(ExePath() + "Export") < 0
 CreateDirectory(ExePath() + "Export")
EndIf

Declare SetGadgetStates()
Declare GetSelectetListAmount(GadgetID.l)
Declare RemoveSelectetItems()
Declare RefreshFileList()
Declare RenameEvent()
Declare CheckPattern(String.s)

Procedure CenterWindowOnMainWindow(MainWindow.l, ChildrenWindow.l)
 If IsWindow(MainWindow) <> 0 And IsWindow(ChildrenWindow) <> 0
  PosX = WindowX(MainWindow) + (WindowWidth(MainWindow)/2) - (WindowWidth(ChildrenWindow)/2)
  PosY = WindowY(MainWindow) + (WindowHeight(MainWindow)/2) - (WindowHeight(ChildrenWindow)/2)
  ResizeWindow(ChildrenWindow, PosX, PosY, #PB_Ignore, #PB_Ignore)
  ProcedureReturn #True
 Else
  ProcedureReturn #False
 EndIf
EndProcedure

Procedure RedrawWinOrGadget(hWin)
 If hWin
  InvalidateRect_(hWin,0,1)
  UpdateWindow_(hWin)
 EndIf
EndProcedure

Procedure AddTextItem(GadgetID.l, Text.s)
 AddGadgetItem(#Stg_Log, -1, Text)
EndProcedure

Procedure.s ConvertNumber(Zahl, Stellen)
 String.s = Str(Zahl)
 ZeroAmount.l = Stellen - Len(String)
 If ZeroAmount > 0
  For a = 1 To ZeroAmount
   String = "0" + String
  Next
  ProcedureReturn String
 Else
  ProcedureReturn String
 EndIf
EndProcedure

Procedure SetGadgetStates()
 If GetGadgetState(#ChkBox_CreateDate) = #False
  DisableGadget(#Cmb_CreateDate, #True)
 Else
  DisableGadget(#Cmb_CreateDate, #False)
 EndIf
 If GetGadgetState(#ChkBox_LastAcces) = #False
  DisableGadget(#Cmb_LastAcces, #True)
 Else
  DisableGadget(#Cmb_LastAcces, #False)
 EndIf
 If GetGadgetState(#ChkBox_LastEdit) = #False
  DisableGadget(#Cmb_LastEdit, #True)
 Else
  DisableGadget(#Cmb_LastEdit, #False)
 EndIf
 
 If GetGadgetState(#Rad_Typ_Copy) = #True
  DisableGadget(#Btn_Copyfolder, #False)
 Else
  DisableGadget(#Btn_Copyfolder, #True)
 EndIf

 If GetGadgetText(#Stg_Replace_1) = ""
  DisableGadget(#Stg_With_1, #True)
 Else
  DisableGadget(#Stg_With_1, #False)
 EndIf
 If GetGadgetText(#Stg_Replace_2) = ""
  DisableGadget(#Stg_With_2, #True)
 Else
  DisableGadget(#Stg_With_2, #False)
 EndIf
 
 If GetGadgetState(#ChkBox_RemoveMisc) = #True
  DisableGadget(#Stg_RemoveMisc_1, #False)
  DisableGadget(#Stg_RemoveMisc_2, #False)
  DisableGadget(#Stg_RemoveMisc_3, #False)
  DisableGadget(#Stg_RemoveMisc_4, #False)
  DisableGadget(#Stg_RemoveMisc_5, #False)
  DisableGadget(#Stg_RemoveMisc_6, #False)
 Else
  DisableGadget(#Stg_RemoveMisc_1, #True)
  DisableGadget(#Stg_RemoveMisc_2, #True)
  DisableGadget(#Stg_RemoveMisc_3, #True)
  DisableGadget(#Stg_RemoveMisc_4, #True)
  DisableGadget(#Stg_RemoveMisc_5, #True)
  DisableGadget(#Stg_RemoveMisc_6, #True)
 EndIf
 
 If CountList(FileList()) > 0
  DisableGadget(#Btn_ClearFiles, #False)
 Else
  DisableGadget(#Btn_ClearFiles, #True)
 EndIf
 If GetSelectetListAmount(#LstVw_Filelist) > 0
  DisableGadget(#Btn_RemoveFiles, #False)
 Else
  DisableGadget(#Btn_RemoveFiles, #True)
 EndIf
 
 If CountGadgetItems(#LstVw_Filelist) = 0 Or GetGadgetText(#Stg_Name) = ""
  DisableGadget(#Btn_Start, #True)
 Else
  DisableGadget(#Btn_Start, #False)
 EndIf
 
 If GetGadgetState(#ExpLst_SearchFile) = -1
  DisableGadget(#Btn_AddFile, #True)
 Else
  DisableGadget(#Btn_AddFile, #False)
 EndIf
EndProcedure

Procedure.l GetSelectetListAmount(GadgetID.l)
 Amount.l = 0
 ListAmount.l = CountGadgetItems(GadgetID) - 1
 For a = 0 To ListAmount
  If GetGadgetItemState(GadgetID, a) = #PB_ListIcon_Selected
   Amount + 1
  EndIf
 Next
 ProcedureReturn Amount
EndProcedure

Procedure RemoveSelectetItems()
 ClearList(DelPos())
 ;Einträge die entfernt werden sollen speichern..
 ForEach FileList()
  If GetGadgetItemState(#LstVw_Filelist, ListIndex(FileList())) = #True
   AddElement(DelPos())
   DelPos() = ListIndex(FileList())
  EndIf
 Next
 ;Einträge entfernen..
 tp.l = -1
 ForEach DelPos()
  tp + 1
  SelectElement(FileList(), DelPos() - tp)
  DeleteElement(FileList())
 Next
EndProcedure

Procedure RefreshFileList()
 ClearGadgetItemList(#LstVw_Filelist)
 ForEach FileList()
  If GetGadgetState(#ChkBox_ShowPath) = #False
   AddGadgetItem(#LstVw_Filelist, -1, GetFilePart(FileList()))
  Else
   AddGadgetItem(#LstVw_Filelist, -1, FileList())
  EndIf
 Next
 SetGadgetText(#Txt_Filelist, Str(CountList(FileList())) + " Dateien")
EndProcedure

Procedure RenameEvent()
 Entr.s = ""
 Name.s = ""
 Exts.s = ""
 Path.s = ""
 StNr.l = Val(GetGadgetText(#Stg_NrStart))
 ErNr.l = Val(GetGadgetText(#Stg_NrFactor))
 AkNr.l = StNr - ErNr
 Prgs.l = 0
 Prga.l = 0
 Amon.l = CountList(FileList())
 ErCo.l = 0
 Resu.l = 0
 Attr.l = 0
 
 If GetGadgetState(#ChkBox_Normal) = #True
  Attr + #PB_FileSystem_Normal
 EndIf
 If GetGadgetState(#ChkBox_ReadOnly) = #True
  Attr + #PB_FileSystem_ReadOnly
 EndIf
 If GetGadgetState(#ChkBox_Hide) = #True
  Attr + #PB_FileSystem_Hidden
 EndIf
 If GetGadgetState(#ChkBox_System) = #True
  Attr + #PB_FileSystem_System
 EndIf

 ForEach FileList()
  If RenameEventT = #True ;Falls Stop gedrückt wird!
   Break
  EndIf
  Prga + 1                                   ;Prozessnummer
  Prgs = (Prga * 100) / Amon                 ;Aktueller Fortschrittprozent
  AkNr + ErNr                                ;Aktuelle Nummer (Nummerierung!)
  Entr = GetGadgetText(#Stg_Name)            ;Aktueller Mustername
  Exts = LCase(GetExtensionPart(FileList())) ;Erweiterung in Kleinbuchstaben
  Name = Left(GetFilePart(FileList()), Len(GetFilePart(FileList())) - (Len(Exts) + 1)) ;Name

  ;*********** Dateiname bearbeiten ********************
  ;Text ersetzen
  If GetGadgetText(#Stg_Replace_1) <> ""
   If GetGadgetState(#ChkBox_IgnoreReplace_1) = #True
    Name = ReplaceString(Name, GetGadgetText(#Stg_Replace_1), GetGadgetText(#Stg_With_1), 1)
   Else
    Name = ReplaceString(Name, GetGadgetText(#Stg_Replace_1), GetGadgetText(#Stg_With_1))
   EndIf
  EndIf
  If GetGadgetText(#Stg_Replace_2) <> ""
   If GetGadgetState(#ChkBox_IgnoreReplace_2) = #True
    Name = ReplaceString(Name, GetGadgetText(#Stg_Replace_2), GetGadgetText(#Stg_With_2), 1)
   Else
    Name = ReplaceString(Name, GetGadgetText(#Stg_Replace_2), GetGadgetText(#Stg_With_2))
   EndIf
  EndIf
  ;Zahlen entfernen
  If GetGadgetState(#ChkBox_RemoveNummeric) = #True
   Name = RemoveString(Name, "1")
   Name = RemoveString(Name, "2")
   Name = RemoveString(Name, "3")
   Name = RemoveString(Name, "4")
   Name = RemoveString(Name, "5")
   Name = RemoveString(Name, "6")
   Name = RemoveString(Name, "7")
   Name = RemoveString(Name, "8")
   Name = RemoveString(Name, "9")
   Name = RemoveString(Name, "0")
  EndIf
  ;Leerzeichen entfernen
  If GetGadgetState(#ChkBox_RemoveSpaces) = #True
   Name = RemoveString(Name, " ")
  EndIf
  ;Sonstiges entfernen
  If GetGadgetState(#ChkBox_RemoveMisc) = #True
   If GetGadgetText(#Stg_RemoveMisc_1) <> ""
    Name = RemoveString(Name, GetGadgetText(#Stg_RemoveMisc_1))
   EndIf
   If GetGadgetText(#Stg_RemoveMisc_2) <> ""
    Name = RemoveString(Name, GetGadgetText(#Stg_RemoveMisc_2))
   EndIf
   If GetGadgetText(#Stg_RemoveMisc_3) <> ""
    Name = RemoveString(Name, GetGadgetText(#Stg_RemoveMisc_3))
   EndIf
   If GetGadgetText(#Stg_RemoveMisc_4) <> ""
    Name = RemoveString(Name, GetGadgetText(#Stg_RemoveMisc_4))
   EndIf
   If GetGadgetText(#Stg_RemoveMisc_5) <> ""
    Name = RemoveString(Name, GetGadgetText(#Stg_RemoveMisc_5))
   EndIf
   If GetGadgetText(#Stg_RemoveMisc_6) <> ""
    Name = RemoveString(Name, GetGadgetText(#Stg_RemoveMisc_6))
   EndIf
  EndIf
  
  ;*********** Mustername bearbeiten *******************
  ;Nicht zugelassene Dateisystemzeichen entfernen
  Entr = RemoveString(Entr, "\")
  Entr = RemoveString(Entr, "/")
  Entr = RemoveString(Entr, ":")
  Entr = RemoveString(Entr, "*")
  Entr = RemoveString(Entr, "?")
  Entr = RemoveString(Entr, Chr(34))
  Entr = RemoveString(Entr, "<")
  Entr = RemoveString(Entr, ">")
  Entr = RemoveString(Entr, "|")
  ;Leerzeichen am Anfang und ende entfernen
  Entr = Trim(Entr)
  ;Nummerierung
  Entr = ReplaceString(Entr, "$Z", ConvertNumber(AkNr, GetGadgetState(#Spn_ZeroAmount)))
  ;Dateiname
  Entr = ReplaceString(Entr, "$N", Name)
  ;Erweiterung
  Entr = ReplaceString(Entr, "$E", Exts)
  ;Dateigröße
  Entr = ReplaceString(Entr, "$S", Str(FileSize((FileList()))))
  ;Stunde
  Entr = ReplaceString(Entr, "%U", FormatDate("%hh", Date()))
  ;Minute
  Entr = ReplaceString(Entr, "%M", FormatDate("%ii", Date()))
  ;Sekunde
  Entr = ReplaceString(Entr, "%S", FormatDate("%ss", Date()))
  ;Tag
  Entr = ReplaceString(Entr, "%T", FormatDate("%dd", Date()))
  ;Monat
  Entr = ReplaceString(Entr, "%O", FormatDate("%mm", Date()))
  ;Jahr
  Entr = ReplaceString(Entr, "%J", FormatDate("%yy", Date()))
  ;Jahr, 4 Stellig
  Entr = ReplaceString(Entr, "%j", FormatDate("%yyyy", Date()))

  Path = "" ;sicherhaltshalber zurücksetzen!
  
  ;Kopieren
  If GetGadgetState(#Rad_Typ_Copy) = #True
   ;Überschreiben.. Frage evtl.
   If GetGadgetState(#ChkBox_OverwriteMsg) = #False And FileSize(SaveFolder + Entr + "." + Exts) > -1
    If MessageRequester(#Programmname, "Die Datei " + SaveFolder + Entr + "." + Exts + " existiert bereits," + #nL + "soll Sie überschrieben werden?", #MB_YESNO|#MB_ICONQUESTION|#MB_DEFBUTTON2) = #IDYES
     AddTextItem(#Stg_Log, "Überschreibe " + SaveFolder + Entr + "." + Exts + " mit " + FileList())
     If CopyFile(FileList(), SaveFolder + Entr + "." + Exts) = #False
      ErCo + 1
      AddTextItem(#Stg_Log, "Fehlgeschlagen.")
     Else
      Path = SaveFolder + Entr + "." + Exts
     EndIf
    Else
     AddTextItem(#Stg_Log, "Datei " + SaveFolder + Entr + "." + Exts + " exisitiert bereits, abgebrochen durch Benutzer.")
    EndIf
   Else
    If FileSize(SaveFolder + Entr + "." + Exts) > -1
     AddTextItem(#Stg_Log, "Überschreibe Datei " + SaveFolder + Entr + "." + Exts + " mit " + FileList())
    Else
     AddTextItem(#Stg_Log, "Kopiere " + FileList() + " nach " + SaveFolder + Entr + "." + Exts)
    EndIf
    If CopyFile(FileList(), SaveFolder + Entr + "." + Exts) = #False
     ErCo + 1
     AddTextItem(#Stg_Log, "Fehlgeschlagen")
    Else
     Path = SaveFolder + Entr + "." + Exts
    EndIf
   EndIf
  EndIf
  
  ;Umbenennen
  If GetGadgetState(#Rad_Typ_RenameOnly) = #True
   AddTextItem(#Stg_Log, "Datei " + FileList() + " wird umbenannt in " + GetPathPart(FileList()) + Entr + "." + Exts)
   If RenameFile(FileList(), GetPathPart(FileList()) + Entr + "." + Exts) = #False
    ErCo + 1
    AddTextItem(#Stg_Log, "Umbenennen fehlgeschlagen.")
   Else
    Path = GetPathPart(FileList()) + Entr + "." + Exts
   EndIf
  EndIf

  If Path <> "" ;falls die datei nicht umbenannt/kopiert werden konnte!
   ;Datum ändern
   If GetGadgetState(#ChkBox_CreateDate) = #True
    If SetFileDate(Path, #PB_Date_Created, GetGadgetState(#Cmb_CreateDate)) = #False
     ErCo + 1
     AddTextItem(#Stg_Log, "Erstelldatum konne nicht geändert werden.")
    EndIf
   EndIf
   If GetGadgetState(#ChkBox_LastEdit) = #True
    If SetFileDate(Path, #PB_Date_Modified, GetGadgetState(#Cmb_LastEdit)) = #False
     ErCo + 1
     AddTextItem(#Stg_Log, "Datum der letzten Änderung konne nicht geändert werden.")
    EndIf
   EndIf
   If GetGadgetState(#ChkBox_LastAcces) = #True
    If SetFileDate(Path, #PB_Date_Accessed, GetGadgetState(#Cmb_LastAcces)) = #False
     ErCo + 1
     AddTextItem(#Stg_Log, "Datum der letzten veränderung konnte nicht verändert werden.")
    EndIf
   EndIf
   ;Attribute ändern
   If Attr <> 0
    If SetFileAttributes(Path, Attr) = #False
     ErCo + 1
     AddTextItem(#Stg_Log, "Dateiattribute konne nicht geändert werden.")
    EndIf
   EndIf
  EndIf
  
  ;Fortschritt anzeigen
  SetGadgetState(#PrgBar_RenameEvent, Prgs)
  SetGadgetText(#Txt_RenameEvent, Str(Prga) + "/" + Str(Amon))
 Next

 ;Protokol vervollständigen
 AddTextItem(#Stg_Log, "")
 AddTextItem(#Stg_Log, "Fehler aufgetretten: " + Str(ErCo))
 SendMessage_(GadgetID(#Stg_Log),#EM_SETSEL,-1,-1)

 ;Thread beenden ect..
 DisableGadget(#Btn_KopieLog, #False)
 DisableGadget(#Btn_SaveLog, #False)
 DisableGadget(#Btn_CloseLog, #False)
 DisableGadget(#Btn_OpenFolder, #False)
 DisableGadget(#Btn_StopRenameEvent, #True) 
 SetGadgetText(#Txt_RenameEvent, "0/0")
 SetGadgetState(#PrgBar_RenameEvent, 0)
 RenameEventT = #False
 MessageBeep_(#MB_ICONINFORMATION)
EndProcedure

Procedure StartRenameEvent()
 SetGadgetText(#Stg_Log, "")
 DisableGadget(#Btn_KopieLog, #True)
 DisableGadget(#Btn_SaveLog, #True)
 DisableGadget(#Btn_CloseLog, #True)
 DisableGadget(#Btn_OpenFolder, #True)
 DisableGadget(#Btn_StopRenameEvent, #False)
 RenameEventT = #False
 DisableWindow(#Win_Main, #True)
 CenterWindowOnMainWindow(#Win_Main, #Win_Log)
 HideWindow(#Win_Main, #True)
 HideWindow(#Win_Log, #False)
 thread = CreateThread(@RenameEvent(), 0)
EndProcedure

Procedure.l CheckPattern(String.s)
 If CountString(String, "/") > 0 Or CountString(String, "\") > 0 Or CountString(String, "|") > 0 Or CountString(String, ":") > 0 Or CountString(String, Chr(34)) > 0
  ProcedureReturn #False
 Else
  ProcedureReturn #True
 EndIf
EndProcedure

;- Startereignisse
AddKeyboardShortcut(#Win_AddFile, #PB_Shortcut_Return, #EnterFileTyp)
SetGadgetStates()
HideWindow(#Win_Main, #False)

;- Programmschleife
Repeat
 Event = WindowEvent()
 If Event

  ;- MenuEvents
  If Event = #PB_Event_Menu
   Select EventMenu()
    Case #Ansicht_Symbole
     ChangeListIconGadgetDisplay(#ExpLst_SearchFile, #PB_ListIcon_LargeIcon)
    
    Case #Ansicht_List
     ChangeListIconGadgetDisplay(#ExpLst_SearchFile, #PB_ListIcon_List)
    
    Case #Ansicht_Details
     ChangeListIconGadgetDisplay(#ExpLst_SearchFile, #PB_ListIcon_Report)
     
    Case #EnterFileTyp
     If CheckPattern(GetGadgetText(#Cmb_FileTyp)) = #True
      SetGadgetText(#ExpLst_SearchFile, GetGadgetText(#Cmb_SearIn) + GetGadgetText(#Cmb_FileTyp))
     EndIf
   
   EndSelect
  EndIf
  
  ;- GadgetEvents
  If Event = #PB_Event_Gadget
   SetGadgetStates()
   
   Select EventGadget()
    Case #Btn_MusterInfo
     String$ = "Der Dateiname darf aus allen zulässigen Zeichen bestehen." + #nL
     String$ + "Ausserdem können Sie die unten stehenden Zeichenfolgen benutzen," + #nL
     String$ + "die mit der jeweiligen Option ausgetauscht werden." + #nL + #nL
     String$ + "Nicht zugelassene Zeichen: \ / : * ? " + Chr(34) + " < > |" + #nL + #nL
     String$ + "Mögliche Zeichenfolgen:" + #nL + #nL
     String$ + "$Z (Nummerierung)" + #nL
     String$ + "$N (Dateiname)" + #nL
     String$ + "$E (Erweiterung)" + #nL
     String$ + "$S (Dateigröße)" + #nL
     String$ + "%U (Stunde)" + #nL
     String$ + "%M (Minute)" + #nL
     String$ + "%S (Sekunde)" + #nL
     String$ + "%T (Tag)" + #nL
     String$ + "%O (Monat)" + #nL
     String$ + "%J (Jahr)" + #nL
     String$ + "%j (Jahr, Vierstellig)" + #nL + #nL
     
     String$ + "%x (entfernt jeweils ein Zeichen vom Anfang des Dateinamens)" + #nL
     String$ + "%y (entfernt jeweils ein Zeichen vom Ende des Dateinamens)" + #nL
     
     MessageRequester(#Programmname, String$, #MB_OK|#MB_ICONINFORMATION)

    Case #Btn_Start
     If CheckFilename(GetGadgetText(#Stg_Name)) = #False
      If MessageRequester(#Programmname, "Ein Dateiname darf keines der folgenden Zeichen enthalten:" + #nL + #nL + "\ / : * ? " + Chr(34) + " < > |" + #nL + "" + #nL + "Ungültige Zeichen werden automatisch entfernt." + #nL + "Möchten Sie fortfahren?", #MB_YESNO|#MB_ICONERROR) = #IDYES
       StartRenameEvent()
      EndIf
     Else
      StartRenameEvent()
     EndIf
         
    Case #Btn_Info
     MessageRequester("Informationen", #Programmname + " " + #Programmvers + #nL + #nL + "Dieses Programm ist Freeware und darf somit" + #nL + "kostenlos verwendet werden." + #nL + #nL + "Copyright©Kai Gartenschläger, 2006" + #nL, #MB_OK|#MB_ICONINFORMATION)
    
    Case #Btn_AddFiles
     CenterWindowOnMainWindow(#Win_Main, #Win_AddFile)
     DisableWindow(#Win_Main, #True)
     HideWindow(#Win_AddFile, #False)
    
    Case #Btn_RemoveFiles
     If GetSelectetListAmount(#LstVw_Filelist) > 0
      RemoveSelectetItems()
      RefreshFileList()
      SetGadgetStates()
     EndIf
    
    Case #Btn_ClearFiles
     ClearList(FileList())
     ClearGadgetItemList(#LstVw_Filelist)
     SetGadgetText(#Txt_Filelist, "0 Dateien")
     SetGadgetStates()
    
    Case #Btn_Copyfolder
     SavePath$ = PathRequester("Speicherordner:", SaveFolder)
     If SavePath$ <> ""
      SaveFolder = SavePath$
     EndIf
    
    Case #Cmb_SearIn
     If GetGadgetText(#Cmb_SearIn) <> GetGadgetText(#ExpLst_SearchFile)
      SearchFolder = GetGadgetText(#Cmb_SearIn)
      SetGadgetText(#ExpLst_SearchFile, SearchFolder)
     EndIf
    
    Case #ExpLst_SearchFile
     If GetGadgetText(#ExpLst_SearchFile) <> GetGadgetText(#Cmb_SearIn)
      SearchFolder = GetGadgetText(#ExpLst_SearchFile)
      SetGadgetText(#Cmb_SearIn, SearchFolder)
     EndIf
    
    Case #Btn_AddFile
     If GetGadgetState(#ExpLst_SearchFile) <> -1
      Files$ = ""
      For a = 0 To CountGadgetItems(#ExpLst_SearchFile) - 1
       If GetGadgetItemState(#ExpLst_SearchFile, a) = #PB_Explorer_File | #PB_Explorer_Selected
        Files$ + GetGadgetItemText(#ExpLst_SearchFile, a, 0) + Chr(34)
       EndIf
      Next
      For a = 1 To CountString(Files$, Chr(34))
       AddElement(FileList())
       FileList() = GetGadgetText(#ExpLst_SearchFile) + StringField(Files$, a, Chr(34))
      Next
      RefreshFileList()
      SetGadgetStates()
      HideWindow(#Win_AddFile, #True)
      SetGadgetText(#ExpLst_SearchFile, SearchFolder)
      SetGadgetText(#Cmb_SearIn, SearchFolder)
      DisableWindow(#Win_Main, #False)
     EndIf

    Case #Btn_AddFileCancel
     HideWindow(#Win_AddFile, #True)
     SetGadgetText(#ExpLst_SearchFile, SearchFolder)
     SetGadgetText(#Cmb_SearIn, SearchFolder)
     DisableWindow(#Win_Main, #False)

    Case #ChkBox_ShowPath
     If CountGadgetItems(#LstVw_Filelist) > 0
      RefreshFileList()
     EndIf
    
    Case #Btn_Help
     If FileSize(ExePath() + "Renamer.chm") > 0
      OpenHelp(ExePath() + "Renamer.chm", "")
     Else
      MessageRequester("Fehler", "Die Hilfedatei konnte nicht gefunden werden.", #MB_OK|#MB_ICONERROR)
     EndIf
    
    Case #Btn_KopieLog
     SetClipboardText(GetGadgetText(#Stg_Log))
     
    Case #Btn_SaveLog
     SavePath$ = SaveFileRequester("Log Datei speichern unter", ExePath(), "Textdatei (*.TXT)|*.TXT|Alle Dateien (*.*)|*.*", 0)
     If SavePath$ <> ""
      OpenFile = CreateFile(#PB_Any, SavePath$)
       WriteString(OpenFile, GetGadgetText(#Stg_Log))
      CloseFile(OpenFile)
     EndIf
    
    Case #Btn_StopRenameEvent
     RenameEventT = #True
    
    Case #Btn_CloseLog
     ClearList(FileList())
     SetGadgetText(#Txt_Filelist, "0 Dateien")
     ClearGadgetItemList(#LstVw_Filelist)
     SetGadgetStates()
     DisableWindow(#Win_Main, #False)
     HideWindow(#Win_Log, #True)
     HideWindow(#Win_Main, #False)
    
    Case #Btn_OpenFolder
     RunProgram("explorer.exe",SaveFolder,"")
     
    Case #Btn_Ansicht
     DisplayPopupMenu(#Ansicht, WindowID(#Win_AddFile), WindowX(#Win_AddFile) + GadgetX(#Btn_Ansicht) + 8, WindowY(#Win_AddFile) + GadgetY(#Btn_Ansicht) + 35)
     
    Case #Btn_ExportFolder
     If GetGadgetText(#ExpLst_SearchFile) <> ExePath() + "Export\"
      SetGadgetText(#ExpLst_SearchFile, ExePath() + "Export\")
      SetGadgetText(#Cmb_SearIn, ExePath() + "Export\")
     EndIf

   EndSelect
  EndIf

  If Event = #PB_Event_SizeWindow
   Select EventWindow()
    Case #Win_AddFile
     ResizeGadget(#Cmb_SearIn, #PB_Ignore, #PB_Ignore, WindowWidth(#Win_AddFile) - 170, #PB_Ignore)
     ResizeGadget(#ExpLst_SearchFile, #PB_Ignore, #PB_Ignore, WindowWidth(#Win_AddFile) - 10, WindowHeight(#Win_AddFile) - 62)
     ResizeGadget(#Btn_AddFile, WindowWidth(#Win_AddFile) - 150, WindowHeight(#Win_AddFile) - 27, #PB_Ignore, #PB_Ignore)
     ResizeGadget(#Btn_AddFileCancel, WindowWidth(#Win_AddFile) - 75, WindowHeight(#Win_AddFile) - 27, #PB_Ignore, #PB_Ignore)
     ResizeGadget(#Btn_Ansicht, WindowWidth(#Win_AddFile) - 107, #PB_Ignore, #PB_Ignore, #PB_Ignore)
     ResizeGadget(#Btn_ExportFolder, WindowWidth(#Win_AddFile) - 55, #PB_Ignore, #PB_Ignore, #PB_Ignore)
     ResizeGadget(#Cmb_FileTyp, #PB_Ignore, WindowHeight(#Win_AddFile) - 27, WindowWidth(#Win_AddFile) - 210, #PB_Ignore)
     ResizeGadget(#Txt_FileTyp, #PB_Ignore, WindowHeight(#Win_AddFile) - 24, #PB_Ignore, #PB_Ignore)
    Case #Win_Log
     ResizeGadget(#Stg_Log, #PB_Ignore, #PB_Ignore, WindowWidth(#Win_Log) - 108, WindowHeight(#Win_Log) - 28)
     ResizeGadget(#Btn_SaveLog, WindowWidth(#Win_Log) - 103, WindowHeight(#Win_Log) - 208, #PB_Ignore, #PB_Ignore)
     ResizeGadget(#Btn_KopieLog, WindowWidth(#Win_Log) - 103, WindowHeight(#Win_Log) - 233, #PB_Ignore, #PB_Ignore)
     ResizeGadget(#Btn_StopRenameEvent, WindowWidth(#Win_Log) - 103, WindowHeight(#Win_Log) - 48, #PB_Ignore, #PB_Ignore)
     ResizeGadget(#Btn_OpenFolder, WindowWidth(#Win_Log) - 103, WindowHeight(#Win_Log) - 183, #PB_Ignore, #PB_Ignore)
     ResizeGadget(#Btn_CloseLog, WindowWidth(#Win_Log) - 103, WindowHeight(#Win_Log) - 23, #PB_Ignore, #PB_Ignore)
     ResizeGadget(#PrgBar_RenameEvent, #PB_Ignore, WindowHeight(#Win_Log) - 13, WindowWidth(#Win_Log) - 108, #PB_Ignore)
     ResizeGadget(#Txt_RenameEvent, #PB_Ignore, WindowHeight(#Win_Log) - 28, WindowWidth(#Win_Log) - 108, #PB_Ignore)
   EndSelect
  EndIf

  If Event = #PB_Event_CloseWindow
   Select EventWindow()
    Case #Win_Main
     End
    Case #Win_AddFile
     HideWindow(#Win_AddFile, #True)
     DisableWindow(#Win_Main, #False)
    Case #Win_Log
     HideWindow(#Win_Log, #True)
   EndSelect
  EndIf
 
 Else
  If IsThread(thread) = #False
   Delay(2)
  EndIf
 EndIf
ForEver

CatchError:
 Msg$ = "Ein Fehler ist in Zeile " + Str(GetErrorLineNR()) + " aufgetreten:" + #nL + #nL
 Msg$ + "Beschreibung: " + GetErrorDescription() + #nL + #nL
 Msg$ + "Fehlerereignisse: " + Str(GetErrorCounter()) + #nL + #nL
 Msg$ + "Bitte senden Sie die Fehlermeldung am Programmautor," + #nL
 Msg$ + "um dieses Programm bei der Verbesserung zu unterstützen." + #nL + #nL
 Msg$ + "Das Programm wird beendet." + #nL + #nL
 Msg$ + "Fehlerbeschreibung in der Zwischenablage kopieren?"
 If MessageRequester(#Programmname, Msg$, #MB_YESNO|#MB_ICONERROR) = #IDYES
  SetClipboardText(Msg$)
 EndIf
 End