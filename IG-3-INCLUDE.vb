IG-3:  ACTIVE COMPONENTS OF IG-3 INCLUDE STATEMENT

   If OptionExists("CLCLLM") Or OptionExists("CLRCLL") Or OptionExists("SPCSPC") Or OptionExists("NG") Or OptionExists("NOG") Then
    	Include = False
    Exit Function
    End If
	
	
	 If Parent.Attributes("CTYPE").Value = "AACA-TRAP" Then
	     If OptionExists("SOLLOC") Or OptionExists("NOG") Or Parent.Attributes("SPG").Value =1 Then
	    	Include = False
	    	Exit Function
	    End If
    Else
	     If OptionExists("SOLLOC") Or OptionExists("NOG") Then 'Or Parent.Attributes("SPG").Value =1 Then
	    	Include = False
	    	Exit Function
	    End If
	End If
	
	  Dim OVERRIDE As Boolean
    OVERRIDE = OptionExists("SCCL") Or OptionExists("SCSGH") Or OptionExists("SCSG") Or _
    	OptionExists("SCSBL") Or OptionExists("CLSBL") Or OptionExists("BZSBL")  _
    	Or OptionExists("SCNSBL")  Or OptionExists("SCBSGT") Or _
    	OptionExists("SCBSLB") Or OptionExists("SCLE") Or OptionExists("AZSBL") 
		
		
		If (InStr(1, Parent.Attributes("CTYPE").Value, "GLASS")  <> 0) And Attributes("H").Value <=36 And Not OVERRIDE _
    	And Not OptionExists("SLD") Then
    	Include = True
    	Exit Function
    End If
	
	
	 If OptionExists("SLD") And (((OptionValue("SLD") + Attributes("H").Value) * 0.5 * Attributes("W").Value) <=2592) And Attributes("W").Value <=72 And Not OVERRIDE Then
    	Include = True
    	Exit Function
    End If
	
	
	
    
    If OptionExists("H2") Then
    	If (((Attributes("SGH").Value + Attributes("H").Value) * 0.5 * Attributes("W").Value) <=2592) And Attributes("W").Value <=72 And Not OVERRIDE Then
    		Include = True
    		Exit Function
    	End If
    End If
	
	   If ((Attributes("W").Value <= 34 And Attributes("H").Value <= 76)  Or (InStr(1, Parent.Attributes("CTYPE").Value, "TRAN") _
	    		 Or InStr(1, Parent.Attributes("CTYPE").Value, "GKW")) <> 0) And Not OVERRIDE And Not OptionExists("SPCSPC") Then
	        Include = True
	    Else 
	    	Include = False
	    End If
		
		 If InStr(1,Parent.Attributes("CTYPE").Value, "GKW") Then
        	If OptionExists("KWNOG") Then
        	Include = False
        	End If
IG-3:  ACTIVE COMPONENTS OF IG-3 INCLUDE STATEMENT

   If OptionExists("CLCLLM") Or OptionExists("CLRCLL") Or OptionExists("SPCSPC") Or OptionExists("NG") Or OptionExists("NOG") Then
    	Include = False
    Exit Function
    End If
	
	
	 If Parent.Attributes("CTYPE").Value = "AACA-TRAP" Then
	     If OptionExists("SOLLOC") Or OptionExists("NOG") Or Parent.Attributes("SPG").Value =1 Then
	    	Include = False
	    	Exit Function
	    End If
    Else
	     If OptionExists("SOLLOC") Or OptionExists("NOG") Then 'Or Parent.Attributes("SPG").Value =1 Then
	    	Include = False
	    	Exit Function
	    End If
	End If
	
	  Dim OVERRIDE As Boolean
    OVERRIDE = OptionExists("SCCL") Or OptionExists("SCSGH") Or OptionExists("SCSG") Or _
    	OptionExists("SCSBL") Or OptionExists("CLSBL") Or OptionExists("BZSBL")  _
    	Or OptionExists("SCNSBL")  Or OptionExists("SCBSGT") Or _
    	OptionExists("SCBSLB") Or OptionExists("SCLE") Or OptionExists("AZSBL") 
		
		
		If (InStr(1, Parent.Attributes("CTYPE").Value, "GLASS")  <> 0) And Attributes("H").Value <=36 And Not OVERRIDE _
    	And Not OptionExists("SLD") Then
    	Include = True
    	Exit Function
    End If

	'Adding comments simply to log a commit
											
	
	
	 If OptionExists("SLD") And (((OptionValue("SLD") + Attributes("H").Value) * 0.5 * Attributes("W").Value) <=2592) And Attributes("W").Value <=72 And Not OVERRIDE Then
    	Include = True
    	Exit Function
    End If
    
    If OptionExists("H2") Then
    	If (((Attributes("SGH").Value + Attributes("H").Value) * 0.5 * Attributes("W").Value) <=2592) And Attributes("W").Value <=72 And Not OVERRIDE Then
    		Include = True
    		Exit Function
    	End If
    End If
	
	   If ((Attributes("W").Value <= 34 And Attributes("H").Value <= 76)  Or (InStr(1, Parent.Attributes("CTYPE").Value, "TRAN") _
	    		 Or InStr(1, Parent.Attributes("CTYPE").Value, "GKW")) <> 0) And Not OVERRIDE And Not OptionExists("SPCSPC") Then
	        Include = True
	    Else 
	    	Include = False
	    End If
		
		 If InStr(1,Parent.Attributes("CTYPE").Value, "GKW") Then
        	If OptionExists("KWNOG") Then
        	Include = False
        	End If
        	End If
