Attribute VB_Name = "rout"
Function nvl(abc As Object, rep As String) As String
If IsNull(abc) Then
nvl = rep
Else
nvl = abc
End If
End Function

Function getrcset(dname As String, qcmd As String, convar As ADODB.Connection, force As Boolean) As ADODB.Recordset
Dim Str As String
Set getrcset = New ADODB.Recordset
If convar = "" Or force Then
convar.Provider = "Microsoft.Jet.OLEDB.4.0;Data Source =" & App.Path & "\" & dname & ".mdb;User Id = admin;Password=;"
convar.CursorLocation = adUseClient
convar.Open
End If
getrcset.Open qcmd, convar, adOpenKeyset, adLockReadOnly
End Function

Function fval() As Boolean

fval = True

With RND

If IsEmpty(.invdate) Or IsEmpty(.Eformdated) Or IsEmpty(.GDFormdated) Or IsEmpty(.MRNOdated) _
 Or IsEmpty(.realdate) Or .invno = "" Or .Companycombo.BoundText = "" Or .Bankscombo.BoundText = "" Or _
 .PartyCombo.BoundText = "" Or .Shipmentby.ListIndex = -1 Or RND.rdscombo.ListIndex = -1 Then
 fval = False
End If

'If .Shipmentby.ListIndex = 0 Then
'  If .Seacombo.BoundText = "" Or .cont_no.Text = "" Or .BLNO.Text = "" _
'  Or IsEmpty(.blnodated) Or .scountry.Text = "" Or IsEmpty(.sarrival_date) Then
'  fval = False
'  End If
'Else
'  If .AWBNO.Text = "" Or IsEmpty(.AWBdated) Or .acountry.Text = "" Or .airportcombo.BoundText = "" _
'  Or .Flightno.Text = "" Or IsEmpty(.aArrival_date) Then
'  fval = False
'  End If
'End If

'Or .exrate.Text = "" Or .valpkr.Text = "" Or .fcyvalue.Text = "" Or .curcombo.BoundText = ""

If .EformNo.Text = "" Or .HSCODES.Text = "" _
  Or .shippieces.Text = "" Or .unitsCombo.BoundText = "" Then
      fval = False
End If

If .rdscombo.ListIndex = 0 Then ' Home Textile

If .rds_solid = "" Then
.rds_solid = 0
End If

If .rds_white = "" Then
.rds_white = 0
End If

If .netpkrval.Text = "" Then
.netpkrval = 0
End If

If .rdsamt.Text = "" Then
.rdsamt.Text = 0
End If

If .rdscharges.Text = "" Then
.rdscharges.Text = 0
End If

If .white_weight.Text = "" Then
.white_weight.Text = 0
End If

If .solid_weight.Text = "" Then
.solid_weight.Text = 0
End If


If .total_weight.Text = "" Then
.total_weight.Text = 0
End If


Else  ' Hosiery

If .TOT_WEIGHT.Text = "" Then
.TOT_WEIGHT.Text = 0
End If

If .netvalpkr_h = "" Then
.netvalpkr_h = 0
End If

If .rds_amt_h = "" Then
.rds_amt_h = 0
End If

If .rds_charges_h = "" Then
.rds_charges_h = 0
End If
End If

End With

End Function

Sub obuts()

If RND.querystat Then
  RND.invno.Locked = False
  RND.findbut.Enabled = False
  RND.exitbut.Enabled = False
  RND.savebut.Enabled = False
  RND.Newbut.Enabled = False
  RND.invno.Locked = False
  RND.repbut.Enabled = False
  RND.Delbut.Enabled = False
Else
  RND.findbut.Enabled = True
  RND.exitbut.Enabled = True
  If (RND.chmd) Then
    RND.savebut.Enabled = True
  End If
  RND.Newbut.Enabled = True
  RND.invno.Locked = True
End If

If RND.newentry Then
  RND.invno.Locked = False
  RND.savebut.Caption = "&Save"
  RND.Delbut.Enabled = False
  RND.repbut.Enabled = False
  RND.savebut.Enabled = False
Else
  
  If Not RND.querystat Then
    RND.repbut.Enabled = True
    RND.Delbut.Enabled = True
  End If

End If

End Sub


Sub clearcurrent()

RND.Company.Recordset.MoveLast
RND.Company.Recordset.MoveNext
RND.party.Recordset.MoveLast
RND.party.Recordset.MoveNext
RND.aport.Recordset.MoveLast
RND.aport.Recordset.MoveNext
RND.Sport.Recordset.MoveLast
RND.Sport.Recordset.MoveNext
RND.units.Recordset.MoveLast
RND.units.Recordset.MoveNext
RND.curren.Recordset.MoveLast
RND.curren.Recordset.MoveNext
RND.invno = ""
RND.Companycombo.Text = ""
RND.PartyCombo.Text = ""
RND.Shipmentby.ListIndex = -1
RND.rdscombo.ListIndex = -1
RND.Seacombo.Text = ""
RND.cont_no.Text = ""
RND.BLNO.Text = ""
RND.scountry.Text = ""
RND.airportcombo.Text = ""
RND.AWBNO.Text = ""
RND.Flightno.Text = ""
RND.acountry.Text = ""
RND.EformNo.Text = ""
RND.GDFormNo.Text = ""
RND.mrno.Text = ""
RND.HSCODES.Text = ""
RND.shippieces.Text = ""
RND.unitsCombo.Text = ""
RND.fcyvalue.Text = ""
RND.curcombo.Text = ""
RND.exrate.Text = ""
RND.valpkr.Text = ""
RND.freight.Text = ""
RND.commission.Text = ""
RND.insurance.Text = ""
RND.netpkrval.Text = ""
RND.rdsamt.Text = ""
RND.rdscharges.Text = ""
RND.fdbcno.Text = ""
chcancel

End Sub

Function findinv() As Boolean
With RND
Set RND.Myrset1 = rout.getrcset("RND", "Select * from RND where invoice_no = '" & .invno.Text & "'", RND.Mycon1, False)
findinv = .Myrset1.RecordCount > 0
If findinv Then
  .invdate = .Myrset1!inv_dated
  .Eformdated = .Myrset1!e_form_dated
  .GDFormdated = .Myrset1!Goods_dated
  .MRNOdated = .Myrset1!mr_dated
  .realdate = .Myrset1!realiz_date
  .invno = .Myrset1!invoice_no
  .Companycombo.BoundText = .Myrset1!company_id
  .PartyCombo.BoundText = .Myrset1!party_id
If Not IsNull(.Myrset1!remarks) Then
  .remtext.Text = .Myrset1!remarks
  End If
    If Not IsNull(.Myrset1!bank_id) Then
    .Bankscombo.BoundText = .Myrset1!bank_id
  Else
    .Bankscombo.BoundText = ""
  End If
  ' record set find
Set RND.Myrset2 = rout.getrcset("RND", "Select * from bill where invoice_no = '" & .invno.Text & "'", RND.Mycon2, False)
Set RND.Myrset3 = rout.getrcset("RND", "Select * from ports where destination_id = " & .Myrset2!port_id, RND.Mycon3, False)
  .Shipmentby.ListIndex = .Myrset2!shipby
If IsNull(.Myrset1!rds_type) Then
  .rdscombo.ListIndex = -1
Else
  .rdscombo.ListIndex = .Myrset1!rds_type
End If
If .Myrset2!shipby = 0 Then
  .Seacombo.BoundText = .Myrset2!port_id
  .cont_no.Text = .Myrset2!f_c_no
  .BLNO.Text = .Myrset2!bill_no
  .blnodated = .Myrset2!bill_dated
  .scountry.Text = .Myrset3!country
  .sarrival_date = .Myrset2!Arrival_date
Else
  .AWBNO.Text = .Myrset2!bill_no
  .AWBdated = .Myrset2!bill_dated
  .acountry.Text = .Myrset3!country
  .airportcombo.BoundText = .Myrset2!port_id
  .Flightno.Text = .Myrset2!f_c_no
  .aArrival_date = .Myrset2!Arrival_date
End If

If .Myrset1!rds_type = 0 Then  ' Home
  .netpkrval.Text = .Myrset1!net_pkr
  .rdsamt.Text = .Myrset1!rds_amount   '  rds_solid+rds+white
  .rds_solid = .Myrset1!rds_solid_5    ' also rds 5% in solid
  .rds_white = .Myrset1!rds_white_3    'also rds 3% in white
  .rdscharges.Text = .Myrset1!rds_service_charges
If Not IsNull(.Myrset1!w_weight) Then
.white_weight.Text = .Myrset1!w_weight
Else
.white_weight.Text = 0
End If
If Not IsNull(.Myrset1!s_weight) Then
.solid_weight.Text = .Myrset1!s_weight
Else
.solid_weight.Text = 0
End If
If Not IsNull(.Myrset1!t_weight) Then
.total_weight.Text = .Myrset1!t_weight
Else
.total_weight.Text = 0

End If



Else   ' Hosiery
If Not IsNull(.Myrset1!t_weight) Then
.TOT_WEIGHT.Text = .Myrset1!t_weight
Else
.TOT_WEIGHT.Text = 0

End If
  
  .netvalpkr_h.Text = .Myrset1!net_pkr
  .rds_amt_h.Text = .Myrset1!rds_amount   ' rds amount 6%
  .rds_charges_h.Text = .Myrset1!rds_service_charges
End If
  .EformNo.Text = .Myrset1!e_form_no
  .GDFormNo.Text = .Myrset1!goods_form_no
  .mrno.Text = .Myrset1!Mr_No
  .HSCODES.Text = .Myrset1!Hs_codes
  .shippieces.Text = .Myrset1!Ship_pieces
  .unitsCombo.BoundText = .Myrset1!unit_id
  If Not IsNull(.Myrset1!d_netshipval) Then
    .netvalship.Text = .Myrset1!d_netshipval
  Else
    .netvalship.Text = 0
  End If
  
  If Not IsNull(.Myrset1!d_tdvs) Then
  .tdvs.Text = .Myrset1!d_tdvs
  Else
  .tdvs.Text = 0
  End If
  
  
  If Not IsNull(.Myrset1!d_shortship) Then
  .shortship.Text = .Myrset1!d_shortship
  Else
  .shortship.Text = 0
  End If
  
  If Not IsNull(.Myrset1!d_nongarment) Then
    .nongarment.Text = .Myrset1!d_nongarment
  Else
    .nongarment.Text = 0
  End If
  
  If Not IsNull(.Myrset1!d_bcharges) Then
    .bcharges.Text = .Myrset1!d_bcharges
  Else
    .bcharges.Text = 0
  End If

    
  .fcyvalue.Text = .Myrset1!fcy_value
  
  .curcombo.BoundText = .Myrset1!currency_id
  .exrate.Text = .Myrset1!exrate
  .valpkr.Text = .Myrset1!val_pkr
  .freight.Text = .Myrset1!freight
  .commission.Text = .Myrset1!commission
  .insurance.Text = .Myrset1!insurance
  .fdbcno.Text = .Myrset1!fdbc_no
  .Companycombo_Click (1)
  .PartyCombo_Click (1)
End If
End With
chsaved
End Function
Function RCALC()
  'fcyvalue
On Error Resume Next
With RND
.netvalship = Val(.tdvs.Text) - Val(.shortship.Text)
Dim usconv As Double
usconv = (.valpkr.Text / Val(RND.USD.Recordset!rupee_conv))
'MsgBox RND.USD.Recordset!rupee_conv

If usconv <= 10000 Then
.rds_charges_h = 300
ElseIf (usconv > 10000 And usconv <= 25000) Then
.rds_charges_h = 600
ElseIf (usconv > 25000) Then
.rds_charges_h = 1000
End If
'
.valpkr.Text = CStr(Val(.exrate.Text) * Val(.fcyvalue.Text))

If .rdscombo.ListIndex = 0 Then ' Home Textile
  .netpkrval.Text = Val(.valpkr.Text) - Val(.freight.Text) - Val(.commission) - Val(.insurance.Text) - Val(.nongarment.Text) - Val(.bcharges.Text)
  .total_weight = Val(.white_weight.Text) + Val(.solid_weight)
  .rds_white.Text = (Val(.white_weight.Text) / Val(.total_weight.Text)) * Val(.netpkrval.Text)
  .rds_solid.Text = Val(.netpkrval.Text) - Val(.rds_white.Text)
  .rdsamt.Text = (Val(.rds_white.Text) * 0.03) + (Val(.rds_solid.Text) * 0.05)
Else  ' Hosiery

  .netvalpkr_h.Text = Val(.valpkr.Text) - Val(.freight.Text) - Val(.commission) - Val(.insurance.Text) - Val(.nongarment.Text) - Val(.bcharges.Text)
  .rds_amt_h.Text = Val(.netvalpkr_h.Text) * 0.06
End If
End With
End Function

Sub saveinvoice()
rout.RCALC

With RND

Set .Myrset1 = rout.getrcset("RND", "delete * from RND where invoice_no = '" & .invno.Text & "'", .Mycon1, False)
Set .Myrset2 = rout.getrcset("RND", "delete * from bill where invoice_no = '" & .invno.Text & "'", .Mycon2, False)

'inserting
Dim qstring As String

qstring = "INSERT INTO RND (invoice_no,inv_dated,company_id,party_id,e_form_no,e_form_dated, goods_form_no,Goods_dated,Mr_No,mr_dated,Hs_codes,Ship_pieces,unit_id,currency_id,exrate,fcy_value,val_pkr,freight,insurance,commission,net_pkr,rds_type,rds_white_3,rds_solid_5,rds_amount,rds_service_charges,fdbc_no,realiz_date,remarks,bank_id,w_weight,s_weight,t_weight,d_netshipval,d_tdvs,d_shortship,d_nongarment,d_bcharges)"

If .fcyvalue.Text = "" Then
.fcyvalue.Text = 0
End If
If .curcombo.BoundText = "" Then
.curcombo.BoundText = 0
End If
 If .Seacombo.BoundText = "" Then
 Seacombo.BoundText = 0
  End If
 
If .airportcombo.BoundText = "" Then
.airportcombo.BoundText = 0
End If


If .exrate.Text = "" Then
.exrate.Text = 0
End If
If .valpkr.Text = "" Then
.valpkr.tex t = 0
End If

If .freight.Text = "" Then
.freight = 0
End If
If .commission = "" Then
.commission = 0
End If
If .insurance.Text = "" Then
.insurance.Text = 0
End If
If .nongarment.Text = "" Then
.nongarment.Text = 0
End If
If .netvalpkr_h.Text = "" Then
.netvalpkr_h.Text = 0
End If

If .rds_amt_h.Text = "" Then
.rds_amt_h.Text = 0
End If

If .rds_charges_h.Text = "" Then
.rds_charges_h.Text = 0
End If

If .TOT_WEIGHT.Text = "" Then
.TOT_WEIGHT.Text = 0
End If

If .netpkrval.Text = "" Then
.netpkrval.Text = 0
End If
If .rds_white.Text = "" Then
.rds_white.Text = 0
End If

If .rds_solid.Text = "" Then
.rds_solid.Text = 0
End If

If .rdsamt.Text = "" Then
.rdscharges.Text = 0
End If

If .rdscharges.Text = "" Then
.rdscharges.Text = 0
End If

If .white_weight.Text = "" Then
.white_weight.Text = 0
End If

If .solid_weight.Text = "" Then
.solid_weight.Text = 0
End If

If .bcharges.Text = "" Then
.bcharges.Text = 0
End If

If .airportcombo.BoundText = "" Then
.airportcombo.BoundText = 0
End If

If .Seacombo.BoundText = "" Then
.Seacombo.BoundText = 0
End If

If .rdscombo.ListIndex = 0 Then   ' Home Textile
qstring = qstring + " values (" & "'" & .invno.Text & "','" & .invdate.Value & "'," & .Companycombo.BoundText & "," & .PartyCombo.BoundText & ",'" & .EformNo.Text & "','" & .Eformdated.Value & "','" & .GDFormNo.Text & "','" _
& .GDFormdated.Value & "','" & .mrno.Text & "','" & .MRNOdated.Value & "','" & .HSCODES.Text & "'," & .shippieces.Text & "," & .unitsCombo.BoundText & "," & .curcombo.BoundText & "," & .exrate.Text & "," & .fcyvalue.Text & "," & .valpkr.Text & "," & .freight.Text & "," & .insurance.Text & "," & .commission.Text & "," & .netpkrval.Text & "," & .rdscombo.ListIndex & "," & .rds_white.Text & "," & .rds_solid.Text & "," & .rdsamt.Text & "," & .rdscharges.Text & ",'" & .fdbcno.Text & "','" & .realdate.Value & "','" & .remtext.Text & "'," & .Bankscombo.BoundText & "," & .white_weight.Text & "," & .solid_weight.Text & "," & .total_weight.Text & "," & .netvalship.Text & "," & .tdvs.Text & "," & .shortship.Text & "," & .nongarment.Text & "," & .bcharges.Text & ")"
Else
qstring = qstring + " values (" & "'" & .invno.Text & "','" & .invdate.Value & "'," & .Companycombo.BoundText & "," & .PartyCombo.BoundText & ",'" & .EformNo.Text & "','" & .Eformdated.Value & "','" & .GDFormNo.Text & "','" _
& .GDFormdated.Value & "','" & .mrno.Text & "','" & .MRNOdated.Value & "','" & .HSCODES.Text & "'," & .shippieces.Text & "," & .unitsCombo.BoundText & "," & .curcombo.BoundText & "," & .exrate.Text & "," & .fcyvalue.Text & "," & .valpkr.Text & "," & .freight.Text & "," & .insurance.Text & "," & .commission.Text & "," & .netvalpkr_h.Text & "," & .rdscombo.ListIndex & ",0,0," & .rds_amt_h.Text & "," & .rds_charges_h.Text & ",'" & .fdbcno.Text & "','" & .realdate.Value & "','" & .remtext.Text & "'," & .Bankscombo.BoundText & ",0,0," & .TOT_WEIGHT.Text & "," & .netvalship.Text & "," & .tdvs.Text & "," & .shortship.Text & "," & .nongarment.Text & "," & .bcharges.Text & ")"
End If

Set RND.Myrset1 = rout.getrcset("RND", qstring, RND.Mycon1, False)

'.invdate.Value
If .Shipmentby.ListIndex = 0 Then  ' By Sea
   Set .Myrset2 = rout.getrcset("RND", "INSERT INTO BILL (invoice_no,shipby,bill_no,bill_dated,f_c_no,arrival_date,port_id) VALUES ('" & .invno.Text & "','" & .Shipmentby.ListIndex & "','" & .BLNO.Text & "','" & .blnodated.Value & "','" & .cont_no.Text & "','" & .sarrival_date & "'," & .Seacombo.BoundText & ")", .Mycon2, False)
 
Else   ' By Air
  Set .Myrset2 = rout.getrcset("RND", "INSERT INTO BILL (invoice_no,shipby,bill_no,bill_dated,f_c_no,arrival_date,port_id) VALUES ('" & .invno.Text & "','" & .Shipmentby.ListIndex & "','" & .AWBNO.Text & "','" & .AWBdated.Value & "','" & .Flightno.Text & "','" & .aArrival_date & "'," & .airportcombo.BoundText & ")", .Mycon2, False)
End If
  .Mycon1.Close
  .Mycon2.Close
.RNDa.Refresh
MsgBox "Invoice Saved!", vbInformation, "Information"
End With
chsaved
End Sub

Sub choccur()
RND.edate.Caption = Format(RND.realdate + 85, "Long Date")
If Not RND.querystat Then
RND.chmd = True
RND.savebut.Enabled = True
If Not RND.newentry Then
RND.savebut.Caption = "&Save Changes"
Else
RND.savebut.Caption = "&Save"
End If
End If
End Sub
Sub chsaved()
RND.chmd = False
RND.newentry = False
RND.savebut.Caption = "&Save"
RND.savebut.Enabled = False
rout.obuts
End Sub
Sub chcancel()
RND.chmd = False
RND.savebut.Enabled = False
End Sub


