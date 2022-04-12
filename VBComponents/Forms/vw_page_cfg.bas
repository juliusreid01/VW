Attribute VB_Name = "vw_page_cfg"

' module is used to prevent cluttering vw_frm_cfg.frm with code used to setup the page

Private pOptions as Collection

Public Property Get Options as Collection: Set Options = pOptions: End Property

Public Sub SetupOptions()
  If pOptions is Nothing Then
    Set pOptions = New Collection
    ConfigureOptions
  End If
End Sub

Public Sub InitPageCfg(p as Page, cfg as Boolean)
  Dim o as vw_option_c

  SetupOptions
  ' set this value to prevent prompt from showing again
  pOptions.Item(S_PAGE_CFG).Value = cfg
  pOptions.Item(S_PAGE_CFG).User = p.PageSheet
  ' prevents application to the shape sheet
  If cfg = False Then Exit Sub
  '//TODO check if in debug mode
  p.PageSheet.OpenSheetWindow
  '//TODO. add more options
  pOptions.Item(S_CHILDOFFSET).DefaultData = p.PageSheet
  pOptions.Item(S_ACTIVELOW).DefaultData = p.PageSheet
  pOptions.Item(S_PERIOD).DefaultData = p.PageSheet
  pOptions.Item(S_SKEW).DefaultData = p.PageSheet
  pOptions.Item(S_DELAY).DefaultData = p.PageSheet
  pOptions.Item(S_DUTYCYCLE).DefaultData = p.PageSheet
  pOptions.Item(S_SIGNALSKEW).DefaultData = p.PageSheet
End Sub

'//TODO request config by name but don't overwrite anything
Public Function ReqConfig(Name as String) as vw_option_c
  If pOptions is Nothing Then SetupOptions
End Function

'//TODO. is this correct, what if some other macro does the same thing
Public Sub DestroyPageCfg(p as Page)
  p.PageSheet.DeleteSection visSectionProp
  p.PageSheet.DeleteSection visSectionUser
End Sub

'//TODO. move to seperate file as all the options are in this code???
Private Function NewOption(Name as String, Optional Desc as String = "") as vw_option_c
  Set NewOption = new vw_option_c
  NewOption.Create Name:=Name, Desc:=Desc
End Function

Private Sub SaveOption(opt as vw_option_c)
  pOptions.Add opt, opt.Name
End Sub

' modify this sub to change details and default values
Private Sub ConfigureOptions()
  Dim opt as vw_option_c

  Set opt = NewOption(Name:=S_PAGE_CFG, Desc:="Indicates this page can control shape behaviour")
  SaveOption opt

  Set opt = NewOption(Name:=S_TYPE, Desc:="Indicates the type of shape")
  opt.Value = 0
  SaveOption opt

  Set opt = NewOption(Name:=S_CHILDOFFSET, Desc:="Minimum distance between linked signals")
  opt.Label = "Child Offset"
  opt.TypeInt = visPropTypeNumber
  opt.Format  = "0.00 u"
  opt.Value   = "BlockSizeY"
  SaveOption opt

  Set opt = NewOption(Name:=S_BUSWIDTH, Desc:="Bus Width")
  opt.Label = "Bus Width"
  opt.TypeInt = visPropTypeNumber
  opt.Format  = "0"
  opt.Value   = 8
  SaveOption opt

  Set opt = NewOption(Name:=S_NAME, Desc:="Enter a name for this shape")
  opt.TypeInt = visPropTypeString
  SaveOption opt

  Set opt = NewOption(Name:=S_CLOCK, Desc:="Reference clock")
  opt.TypeInt = visPropTypeString
  opt.Hidden = "STRSAME("""", Prop." & S_CLOCK & ".Format)"
  SaveOption opt

  Set opt = NewOption(Name:=S_SIGNAL, Desc:="Reference signal")
  opt.TypeInt = visPropTypeString
  opt.Hidden = "STRSAME("""", Prop." & S_SIGNAL & ".Format)"
  SaveOption opt

  Set opt = NewOption(Name:=S_ACTIVELOW, Desc:="Polarity")
  opt.Label = "Active Low"
  opt.TypeInt = visPropTypeBool
  opt.Value = False
  SaveOption opt

  Set opt = NewOption(Name:=S_PERIOD, Desc:="Period")
  opt.TypeInt = visPropTypeListNumber
  opt.Format = "0.00 u"
  opt.Value = "BlockSizeX*2"
  SaveOption opt

  Set opt = NewOption(Name:=S_SKEW, Desc:="Percentage of Period/2 to delay edges")
  opt.Label = "Skew %"
  opt.TypeInt = visPropTypeListNumber
  opt.Format = "0.0 u"
  opt.Value = "10 %"
  SaveOption opt

  Set opt = NewOption(Name:=S_DELAY, Desc:="Initial clock delay or signal transition delay")
  opt.TypeInt = visPropTypeListNumber
  opt.Format = "0.000 u"
  opt.Value = "BlockSizeX*0.5"
  SaveOption opt

  Set opt = NewOption(Name:=S_DUTYCYCLE, Desc:="Clock Duty Cycle")
  opt.Label = "Duty Cycle %"
  opt.TypeInt = visPropTypeListNumber
  opt.Format = "0.00 u"
  opt.Value = "50 %"
  SaveOption opt

  Set opt = NewOption(Name:=S_SIGNALSKEW, Desc:="Amount of skew to apply to dependent signals")
  opt.Label = "Signal Skew"
  opt.TypeInt = visPropTypeListNumber
  '//TODO this format is false for page sheet
  opt.Format = "0.000 u"
  opt.Value = VW_0 & "+ (Prop." & S_PERIOD & "* 0.5 * Prop." & S_SKEW & ")"
  SaveOption opt

  '//TODO. we need a better description
  Set opt = NewOption(Name:=S_EVENTTYPE, Desc:="Select an event to add/modify")
  opt.Label = "Event Type"
  opt.TypeInt = visPropTypeListFix
  SaveOption opt

  Set opt = NewOption(Name:=S_EVENTTRIGGER, Desc:="Select a trigger type for the event")
  opt.Label = "Trigger"
  opt.TypeInt = visPropTypeListFix
  opt.Hidden = "STRSAME("""", Prop." & S_EVENTTYPE & ")"
  SaveOption opt

  Set opt = NewOption(Name:=S_EVENTPOSITION, Desc:="Input the positon based on the trigger")
  opt.Label = "Position"
  opt.TypeInt = visPropTypeListFix
  opt.Hidden = "OR(Prop." & S_EVENTTRIGGER & ".Invisible,STRSAME("""", Prop." & S_EVENTTRIGGER & "))"
  SaveOption opt

  Set opt = NewOption(Name:=S_LABELEDGES, Desc:="Select which labels to show on transitions")
  opt.Label = "Label Edges"
  opt.TypeInt = visPropTypeListVar
  opt.Hidden = "STRSAME("""", Prop." & S_LABELEDGES & ".Format)"
  SaveOption opt

  Set opt = NewOption(Name:=S_LABELSIZE, Desc:="Height of the labels")
  opt.Label = "Label Size"
  opt.TypeInt = visPropTypeListNumber
  opt.Format = "0.000 u"
  opt.Value = "BlockSizeX*0.667"
  opt.Hidden = "OR(Prop." & S_LABELEDGES & ".Invisible, STRSAME(""None"", Prop." & S_LABELEDGES & "))"
  SaveOption opt

  Set opt = NewOption(Name:=S_LABELFONT, Desc:="Font Size of the labels")
  opt.Label = "Label Font Pt"
  opt.TypeInt = visPropTypeListNumber
  opt.Format = "0 u"
  opt.Value = "8 pt"
  opt.Hidden = "Prop." & S_LABELSIZE & ".Invisible"
  SaveOption opt

  Set opt = NewOption(Name:=S_NODEFONT, Desc:="Font size of the nodes")
  opt.Label = "Node Font Pt"
  opt.TypeInt = visPropTypeListNumber
  opt.Format = "0 u"
  opt.Value = "6 pt"
  opt.Hidden = True
  SaveOption opt

  Set opt = NewOption(Name:=S_NODESIZEMULT, Desc:="Height of the nodes")
  opt.Label = "Node Size Multiplier"
  opt.TypeInt = visPropTypeListNumber
  opt.Format = "0.00 u"
  opt.Value = "0.5"
  opt.Hidden = True
  SaveOption opt
End Sub