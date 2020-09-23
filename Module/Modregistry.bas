Attribute VB_Name = "Modregistry"
Option Explicit

Sub saving_withTax_rec()

Call SaveSetting("WithHolding Tax", "Exemption", "ExemptVal1", frmWithTax.txtex1.Text)
Call SaveSetting("WithHolding Tax", "Exemption", "ExemptVal2", frmWithTax.txtex2.Text)
Call SaveSetting("WithHolding Tax", "Exemption", "ExemptVal3", frmWithTax.txtex3.Text)
Call SaveSetting("WithHolding Tax", "Exemption", "ExemptVal4", frmWithTax.txtex4.Text)
Call SaveSetting("WithHolding Tax", "Exemption", "ExemptVal5", frmWithTax.txtex5.Text)
Call SaveSetting("WithHolding Tax", "Exemption", "ExemptVal6", frmWithTax.txtex6.Text)
Call SaveSetting("WithHolding Tax", "Exemption", "ExemptVal7", frmWithTax.txtex7.Text)
Call SaveSetting("WithHolding Tax", "Exemption", "ExemptVal8", frmWithTax.txtex8.Text)

Call SaveSetting("WithHolding Tax", "Status", "StatOOPVal1", frmWithTax.txtop1.Text)
Call SaveSetting("WithHolding Tax", "Status", "StatOOPVal2", frmWithTax.txtop2.Text)
Call SaveSetting("WithHolding Tax", "Status", "StatOOPVal3", frmWithTax.txtop3.Text)
Call SaveSetting("WithHolding Tax", "Status", "StatOOPVal4", frmWithTax.txtop4.Text)
Call SaveSetting("WithHolding Tax", "Status", "StatOOPVal5", frmWithTax.txtop5.Text)
Call SaveSetting("WithHolding Tax", "Status", "StatOOPVal6", frmWithTax.txtop6.Text)
Call SaveSetting("WithHolding Tax", "Status", "StatOOPVal7", frmWithTax.txtop7.Text)
Call SaveSetting("WithHolding Tax", "Status", "StatOOPVal8", frmWithTax.txtop8.Text)


End Sub

Public Sub saving_rice_rec()

Call SaveSetting("rice_allowance", "cost", "Rice Code", frmriceallowance.Text1.Text)
Call SaveSetting("rice_allowance", "cost", "Rice Code1", frmriceallowance.Text2.Text)
Call SaveSetting("rice_allowance", "cost", "Rice allowance", frmriceallowance.Text3.Text)
Call SaveSetting("rice_allowance", "cost", "Rice allowance2", frmriceallowance.Text4.Text)

End Sub

Public Sub saving_living_rec()

Call SaveSetting("living_allowance", "cost", "living Code", frmliving.Text1.Text)
Call SaveSetting("living_allowance", "cost", "living Code1", frmliving.Text2.Text)
Call SaveSetting("living_allowance", "cost", "living allowance", frmliving.Text3.Text)
Call SaveSetting("living_allowance", "cost", "living allowance2", frmliving.Text4.Text)

End Sub

Public Sub saving_SSS_Table_rec()

'-// Save Code
Call SaveSetting("SSS Table", "cost", "SSS Code1", frmssstable.txt1.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code2", frmssstable.txt2.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code3", frmssstable.txt3.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code4", frmssstable.txt4.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code5", frmssstable.txt5.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code6", frmssstable.txt6.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code7", frmssstable.txt7.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code8", frmssstable.txt8.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code9", frmssstable.txt9.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code10", frmssstable.txt10.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code11", frmssstable.txt11.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code12", frmssstable.txt12.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code13", frmssstable.txt13.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code14", frmssstable.txt14.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code15", frmssstable.txt15.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code16", frmssstable.txt16.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code17", frmssstable.txt17.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code18", frmssstable.txt18.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code19", frmssstable.txt19.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code20", frmssstable.txt20.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code21", frmssstable.txt21.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code22", frmssstable.txt22.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code23", frmssstable.txt23.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code24", frmssstable.txt24.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code25", frmssstable.txt25.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code26", frmssstable.txt26.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code27", frmssstable.txt27.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code28", frmssstable.txt28.Text)
Call SaveSetting("SSS Table", "cost", "SSS Code29", frmssstable.txt29.Text)

Call SaveSetting("SSS Table", "EShare", "SSS Share1", frmssstable.txtE1.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share2", frmssstable.txtE2.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share3", frmssstable.txtE3.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share4", frmssstable.txtE4.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share5", frmssstable.txtE5.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share6", frmssstable.txtE6.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share7", frmssstable.txtE7.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share8", frmssstable.txtE8.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share9", frmssstable.txtE9.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share10", frmssstable.txtE10.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share11", frmssstable.txtE11.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share12", frmssstable.txtE12.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share13", frmssstable.txtE13.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share14", frmssstable.txtE14.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share15", frmssstable.txtE15.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share16", frmssstable.txtE16.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share17", frmssstable.txtE17.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share18", frmssstable.txtE18.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share19", frmssstable.txtE19.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share20", frmssstable.txtE20.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share21", frmssstable.txtE21.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share22", frmssstable.txtE22.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share23", frmssstable.txtE23.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share24", frmssstable.txtE24.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share25", frmssstable.txtE25.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share26", frmssstable.txtE26.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share27", frmssstable.txtE27.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share28", frmssstable.txtE28.Text)
Call SaveSetting("SSS Table", "EShare", "SSS Share29", frmssstable.txtE29.Text)
 

End Sub

Sub saving_PH_share()

Call SaveSetting("PhilHealth Table", "EShare", "PH Share1", frmPhilHealth.txtE1)
Call SaveSetting("PhilHealth Table", "EShare", "PH Share2", frmPhilHealth.txtE2)
Call SaveSetting("PhilHealth Table", "EShare", "PH Share3", frmPhilHealth.txtE3)
Call SaveSetting("PhilHealth Table", "EShare", "PH Share4", frmPhilHealth.txtE4)
Call SaveSetting("PhilHealth Table", "EShare", "PH Share5", frmPhilHealth.txtE5)
Call SaveSetting("PhilHealth Table", "EShare", "PH Share6", frmPhilHealth.txtE6)
Call SaveSetting("PhilHealth Table", "EShare", "PH Share7", frmPhilHealth.txtE7)
Call SaveSetting("PhilHealth Table", "EShare", "PH Share8", frmPhilHealth.txtE8)
Call SaveSetting("PhilHealth Table", "EShare", "PH Share9", frmPhilHealth.txtE9)
Call SaveSetting("PhilHealth Table", "EShare", "PH Share10", frmPhilHealth.txtE10)
Call SaveSetting("PhilHealth Table", "EShare", "PH Share11", frmPhilHealth.txtE11)
Call SaveSetting("PhilHealth Table", "EShare", "PH Share12", frmPhilHealth.txtE12)
Call SaveSetting("PhilHealth Table", "EShare", "PH Share13", frmPhilHealth.txtE13)
Call SaveSetting("PhilHealth Table", "EShare", "PH Share14", frmPhilHealth.txtE14)
Call SaveSetting("PhilHealth Table", "EShare", "PH Share15", frmPhilHealth.txtE15)
Call SaveSetting("PhilHealth Table", "EShare", "PH Share16", frmPhilHealth.txtE16)
Call SaveSetting("PhilHealth Table", "EShare", "PH Share17", frmPhilHealth.txtE17)

End Sub
