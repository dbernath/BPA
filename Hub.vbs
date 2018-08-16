'******************** Black Pearl Workflow - Development *************************
'******************** Hub ***************************************************
'***************************************************************************

'******************** Black Pearl Workflow - Development *************************
'******************** Hub ***************************************************
'******************** Hub ***************************************************
Function Hub()
Logger.Info "Hub - Start"
Dim strDrawer:						strDrawer = File.Drawer.Name:																														Logger.Info "Hub - Drawer - "& strDrawer
Dim strStatus:						strStatus = Task.GetAttributeObject("ADJUSTMENT_STATUS").Value:													Logger.Info "Hub - Status - " & strStatus
Dim strContact:					strContact = Task.GetAttributeObject("BPA_CONTACT").Value:																Logger.Info "Hub - Contact - " & strContact
Dim strAdjust:						strAdjust = Task.GetAttributeObject("BPA_ADJUSTING_TYPE").Value:													Logger.Info "Hub - Adjusting - " & strAdjust
Dim strIndexing:					strIndexing = Task.GetAttributeObject("BPA_INDEXING_TYPE").Value:													Logger.Info "Hub - Indexing - " &strIndexing
Dim strFile:							strFile = Task.File.FullFileNumber:																													Logger.Info "Hub - File - " & strFile
'Dim strUser:							strUser = Task.FromUser.Name:																													Logger.Info "Hub - From User - " & strUser
Dim strUser:							strUser = "rschwinn":																																		Logger.Info "Hub - From User - " & strUser
Dim strParameter:				strParameter = ""
Dim strDelivery:					strDelivery = Task.GetAttributeObject("DELIVERY_METHOD").Value:													Logger.Info "Hub - Delivery Method - " & strDelivery
Dim strPaymentType:			strPaymentType = Task.GetAttributeObject("TYPE_OF_CHECK_REQUEST").Value:								Logger.Info "Hub - Payment Type - " & strPaymentType
Dim strFlow:							strFlow = ""
Dim strStep:							strStep = ""
Dim strPriority:						strPriority = ""
Dim strTaskDescription:		strTaskDescription = ""
Dim strNote, strInsd, strPolicy, sResult, strSQL, strDB, strAdj

Select Case strStatus
	Case ""
		Logger.Info "Hub - New Work"
		If Left(strFile,3) = "~ER" Then
			Hub = 8 ' Send to Indexing
		Else
			Logger.Info "Hub - New Work with claim number."
			If File.GetAttributeObject("IA/ADJ").Value = "BPA" Then
				Task.Description = "Adjusting Work"
				File.GetAttributeObject("ADJUSTMENT_STATUS").Value = "Correspondence Awaiting Review"
				Find_Contact_Form objContactForm 
				Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_RECEIVED_FROM").Value = Document.GetAttributeObject("ER_FROMADDR").Value
				Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_INSD").Value = objContactForm.Form.GetFieldValue ("//data/ClaimPrimaryEmail")
				Task.AssignedTo = AccountsLookup.GetUser("bpauser")
					Select Case strDrawer
						Case "CPCL"
							Logger.Info "Hub - Indexed - Adjusting Work - CPIC"
							Hub = 5
						Case "SFCL"
							Logger.Info "Hub - Indexed - Adjusting Work - SFIC"
							Hub = 6
						Case Else
							Logger.Info "Hub - Indexed - Adjusting Work - SFPC"
							Hub = 7
					End Select
				strParameter = "[%MESSAGE_TO%]:" & Document.GetAttributeObject("ER_FROMADDR").Value
				Update_CDS_Status File.GetAttributeObject("ADJUSTMENT_STATUS").Value, strParameter
			Else
				Logger.Info "Hub - New Work with claim number but not BPA"
				Hub = 12
			End If
		End If
			
		Exit Function

	Case "New Loss"
		Logger.Info "Hub - New Loss"
'		Task.AssignedTo = AccountsLookup.GetUser("Black Pearl User")
		strSQL = "Select dbo.CDS_Email_Verified ('" + Task.File.FullFileNumber + "')":																Logger.Info "Hub - Indexed - SQL - " & strSQL
		strDB = "PROVIDER=SQLOLEDB;Data Source=ir5dbdev;Initial Catalog=Utility;uid=sa;PWD=cpic0742":						Logger.Info "Hub - Indexed - DB - " & strDB
		 Call_Database strSQL, strDB, sResult
		
		If sResult = "Not Verified" Then
			Select Case strDrawer
				Case "CPCL"
					Logger.Info "Hub - New Loss - CPIC"
					Hub = 1 ' Send to CPIC New Loss
				Case "SFCL"
					Logger.Info "Hub - New Loss - SFIC"
					Hub = 2 ' Send to SFIC New Loss
				Case Else
					Logger.Info "Hub - New Loss - SFPC"
					Hub = 3 ' Send to SFPC New Loss
			End Select
'			strNote =  "BPA email verification link was not verified." ' Duplicate Note as CDS
			Get_Contact_Form_Info
		Else
'			strNote =  "BPA email verification link was verified." ' Duplicate Note as CDS
			Hub = 9
		End If
		File.GetAttributeObject("ADJUSTMENT_STATUS").Value = "New Loss"
		Update_CDS_Status File.GetAttributeObject("ADJUSTMENT_STATUS").Value, strParameter
		File.GetAttributeObject("ADJUSTMENT_STATUS").Value = "Awaiting Estimates"
		Update_CDS_Status File.GetAttributeObject("ADJUSTMENT_STATUS").Value, strParameter
'		AutoNote strDrawer, strFile, strFlow, strStep, strUser, strPriority, strTaskDescription, strNote  ' Duplicate Note as CDS
		Exit Function
	Case "Contact"
		Logger.Info "Hub - Contact"
		Find_Contact_Form objContactForm
		
		strMortgage = 	objContactForm.Form.GetFieldValue ("//data/Mortgage1Name1") & " " & _
									objContactForm.Form.GetFieldValue ("//data/Mortgage1Name2") & " " & _
									objContactForm.Form.GetFieldValue ("//data/Mortgage1Name3") & ", " & _
									objContactForm.Form.GetFieldValue ("//data/Mortgage1Address") & ", " & _
									objContactForm.Form.GetFieldValue ("//data/Mortgage1CityStateZip") & _
									", Loan # " & objContactForm.Form.GetFieldValue ("//data/Mortgage1Loan")

		Dim strEmail:	 			strEmail = Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_INSD").Value:					Logger.info "Hub - Contact Email - " & strEmail
		Dim strPhone:	 		strPhone = Task.GetAttributeObject("BPA_ADJUSTING_HOME_INSD").Value:					Logger.info "Hub - Contact Phone - " & strPhone
		Dim strCell:		 		strCell = Task.GetAttributeObject("BPA_ADJUSTING_CELL_INSD").Value:							Logger.info "Hub - Contact Cell - " & strCell
		Dim strAltName:		strAltName = Task.GetAttributeObject("BPA_ADJUSTING_NAME_ALT").Value:				Logger.info "Hub - Alt Name - " & strAltName
		Dim strAltEmail:		strAltEmail = Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_ALT").Value:				Logger.info "Hub - Alt Email - " & strAltEmail
		Dim strAltHome: 		strAltHome = Task.GetAttributeObject("BPA_ADJUSTING_HOME_ALT").Value:				Logger.info "Hub - Alt Home - " & strAltHome
		Dim strAltCell:		 	strAltCell = Task.GetAttributeObject("BPA_ADJUSTING_CELL_ALT").Value:						Logger.info "Hub - Alt Cell - " & strAltCell
		Dim strMrtg:		 		strMrtg = Task.GetAttributeObject("BPA_MORTGAGE_CORRECTION").Value:					Logger.info "Hub - Contact Mrtg - " & strMrtg
		
		
		Dim strChanges:			strChanges = ""
		Dim strOldEmail:			strOldEmail = objContactForm.Form.GetFieldValue ("//data/ClaimPrimaryEmail"):								Logger.info "Hub - Original Contact Email - " & strOldEmail
		Dim strOldPhone:			strOldPhone = objContactForm.Form.GetFieldValue ("//data/ClaimPrimaryHome"):							Logger.info "Hub - Original Contact Phone - " & strOldPhone
		Dim stOldCell:		 		stOldCell = objContactForm.Form.GetFieldValue ("//data/ClaimPrimaryCell"):										Logger.info "Hub - Contact Cell - " & stOldCell
		Dim strOldAltName:	strOldAltName = objContactForm.Form.GetFieldValue ("//data/ClaimSecondaryName"):					Logger.info "Hub - Alt Name - " & strOldAltName
		Dim strOldAltEmail:		strOldAltEmail = objContactForm.Form.GetFieldValue ("//data/ClaimSecondaryEmail"):						Logger.info "Hub - Alt Email - " & strOldAltEmail
		Dim strOldAltHome: 	strOldAltHome = objContactForm.Form.GetFieldValue ("//data/ClaimSecondaryHome"):					Logger.info "Hub - Alt Home - " & strOldAltHome
		Dim strOldAltCell:		strOldAltCell = objContactForm.Form.GetFieldValue ("//data/ClaimSecondaryCell"):							Logger.info "Hub - Alt Cell - " & strOldAltCell
		Dim strOldMrtg:			strOldMrtg = strMortgage:																																	Logger.info "Hub - Original Contact Mrtg - " & strOldMrtg
		Dim strNoteDrawer

	'Check to see if any fields changed and need to update CDS and POINT	
		If strEmail <> strOldEmail Then
			Logger.Info "Hub - Email Correction"
			strChanges = strChanges & "<p> <b>Original Insured Email Address:</b> " & strOldEmail & "<br/> <b>Revised Insured Email Address:</b> " & strEmail  & "</p>"
			objContactForm.Form.SetFieldValue "//data/ClaimPrimaryEmail", strEmail
		End If
		If strPhone <> strOldPhone Then
			Logger.Info "Hub - Phone Correction"
			strChanges = strChanges & "<p> <b>Original Insured Home Phone:</b> " & strOldPhone & "<br/> <b>Revised Insured Home Phone:</b> " & strPhone  & "</p>"
			objContactForm.Form.SetFieldValue "//data/ClaimPrimaryHome", strPhone
		End If
		If strCell <> stOldCell Then
			Logger.Info "Hub - Cell Correction"
			strChanges = strChanges & "<p> <b>Original Insured Cell Phone:</b> " & stOldCell & "<br/> <b>Revised Insured Cell Phone:</b> " & strCell  & "</p>"
			objContactForm.Form.SetFieldValue "//data/ClaimPrimaryCell", strCell
		End If
		If strAltName <> strOldAltName Then
			Logger.Info "Hub - Alt Name Correction"
			strChanges = strChanges & "<p> <b>Original Alternate Contact Name:</b> " & strOldAltName  & "<br/><b>Revised Alternate Contact Name:</b> " & strAltName & "</p>"
			objContactForm.Form.SetFieldValue "//data/ClaimSecondaryName", strAltName
		End If
		If strAltEmail <> strOldAltEmail Then
			Logger.Info "Hub - Alt Email Correction"
			strChanges = strChanges & "<p> <b>Original Alternate Contact Email:</b> " & strOldAltEmail  & "<br/><b>Revised Alternate Contact Email:</b> " & strAltEmail & "</p>"
			objContactForm.Form.SetFieldValue "//data/ClaimSecondaryEmail", strAltEmail
		End If
		If strAltHome <> strOldAltHome Then
			Logger.Info "Hub - Alt Home Phone Correction"
			strChanges = strChanges & "<p> <b>Original Alternate Contact Home Phone:</b> " & strOldAltHome  & "<br/><b>Revised Alternate Contact Home Phone:</b> " & strAltHome & "</p>"
			objContactForm.Form.SetFieldValue "//data/ClaimSecondaryHome", strAltHome
		End If
		If strAltCell <> strOldAltCell Then
			Logger.Info "Hub - Alt Cell Phone Correction"
			strChanges = strChanges & "<p> <b>Original Alternate Contact Cell Phone:</b> " & strOldAltCell  & "<br/><b>Revised Alternate Contact Cell Phone:</b> " & strAltCell & "</p>"
			objContactForm.Form.SetFieldValue "//data/ClaimSecondaryCell", strAltCell
		End If
		If strMrtg <> strOldMrtg Then
			Logger.Info "Hub - Mortgage Correction"
			strChanges = strChanges & "<p> <b>Original Mortgage:</b> " & strOldMrtg & "<br/> <b>Revised Mortgage:</b> " & strMrtg & "</p>"
		End If

		If strChanges <> "" Then
			Logger.Info "Hub - Correction Start"
			Select Case File.Drawer.Name
				Case "CPCL"
					strNoteDrawer = "CPUW"
				Case "SFCL"
					strNoteDrawer = "SFUW"
				Case Else
					strNoteDrawer = "SPCU"
			End Select		
			Logger.Info "Hub - Note Drawer - " & strNoteDrawer
			strPolicy = objContactForm.Form.GetFieldValue ("//data/PolicyNumber"):			Logger.Info "Hub - File - " & strPolicy
			strFlow = "INFORCE":																										Logger.Info "Hub - File - " & strFlow
			strStep = "S_b66c62a0-d455-4e11-93c3-15afb337309f":										Logger.Info "Hub - File - " & strStep
			strNote = "<p>During a claim call the insured indicated that our records are incorrect and need to be updated with the information below.</p>" & strChanges & _
									"<p>Please review this information and make the appropriate changes. Contact the insured or agent if additional information is needed to make the changes."
			Logger.Info "Hub - Changes - " & strNote
			AutoNote strNoteDrawer, strPolicy, strFlow, strStep, strUser,"4", "BPA Corrections", strNote
			objContactForm.Form.Save
			Update_CDS_Contact Task.File.FullFileNumber, strEmail, strCell, strPhone
			Update_CDS_Alternate_Contact Task.File.FullFileNumber, strAltEmail, strAltHome, strAltCell, strAltName
			strFlow = ""
			strStep = ""
			strNote = ""
		End If
		
		Select Case strContact
			Case "Spoke with insured"
				File.GetAttributeObject("ADJUSTMENT_STATUS").Value = "Contacted - Awaiting Estimates"
				strNote = "Contacted insured. "
				Hub = 9
			Case "Left message for insured"
				File.GetAttributeObject("ADJUSTMENT_STATUS").Value = "Uncontacted - Awaiting Estimates"
				Task.GetAttributeObject("BPA_CONTACT").Value = ""
				Task.DateAvailable = DateAdd("h",4,DateValue(Date()+1)):				Logger.Info "Hub - Uncontacted - Date Available - " & Task.DateAvailable
				Task.Description = "Another contact attempt."
				Select Case strDrawer
					Case "CPCL"
						Logger.Info "Hub - Another contact attempt - CPIC"
						Hub = 1
					Case "SFCL"
						Logger.Info "Hub - Another contact attempt - SFIC"
						Hub = 2
					Case Else
						Logger.Info "Hub - Another contact attempt - SFPC"
						Hub = 3
				End Select
				strNote = "Unable to contact insured. Task has been rescheduled for tomorrow."
			Case "Send to Reassignment"
				File.GetAttributeObject("ADJUSTMENT_STATUS").Value = "Pending Reassignment"
				Task.Description = "Reassign Claim"
				strNote = "The task has been sent to the Reassignment step."
				File.ClearFileMarkById(33462655)'BPA First Follow Up Sent 					Production 59519388			Development 33462655
				File.ClearFileMarkById(33462656)'BPA Second Follow Up Sent				Production 59519450 			Development 33462656
				File.ClearFileMarkById(33462657)'BPA Final Follow Up Sent					Production 59519507			Development 33462657
				File.ClearFileMarkById(33453717)'Black Pearl Adjusting							Production 57201953			Development 33453717
				Hub = 4
			Case "Send to Less than Deductible"
				File.GetAttributeObject("ADJUSTMENT_STATUS").Value = "Less than Deductible"
				strNote = "The task has been sent to the Less than Deductible step."
				Task.Description = "BPA Less than Deductible"
				Hub = 4
			Case "Send to Withdrawl"
				File.GetAttributeObject("ADJUSTMENT_STATUS").Value = "Withdrawn by Insured"
				strNote = "The task has been sent to the Withdrawn step."
				Task.Description = "BPA Withdrawl"
				Hub = 4
		End Select		
		Update_CDS_Status File.GetAttributeObject("ADJUSTMENT_STATUS").Value, strParameter
		AutoNote strDrawer, strFile, strFlow, strStep, strUser, strPriority, strTaskDescription, strNote
		Exit Function
	Case "Indexed"
		Task.Priority = 5
		Select Case strIndexing
			Case "Adjusting Work"
				strSQL = "select Utility.dbo.Lookup_Claim_Prefix_Drawer ('" + File.FullFileNumber + "','" + File.Drawer.Name + "')":									Logger.Info "Hub - Indexed - SQL - " & strSQL
				strDB = "PROVIDER=SQLOLEDB;Data Source=ir5dbdev;Initial Catalog=Utility;uid=sa;PWD=cpic0742":																Logger.Info "Hub - Indexed - DB - " & strDB
				 Call_Database strSQL, strDB, sResult
				If sResult = "Wrong Drawer" Then
					Task.Description = "Please correct the drawer."
					Task.Priority = 1
					Hub = 8
					Exit Function
				End If
				Find_Contact_Form objContactForm 
				Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_RECEIVED_FROM").Value = Document.GetAttributeObject("ER_FROMADDR").Value
				Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_INSD").Value = objContactForm.Form.GetFieldValue ("//data/ClaimPrimaryEmail")
				Task.AssignedTo = AccountsLookup.GetUser("bpauser")
					Select Case strDrawer
						Case "CPCL"
							Logger.Info "Hub - Indexed - Adjusting Work - CPIC"
							Hub = 5
						Case "SFCL"
							Logger.Info "Hub - Indexed - Adjusting Work - SFIC"
							Hub = 6
						Case Else
							Logger.Info "Hub - Indexed - Adjusting Work - SFPC"
							Hub = 7
					End Select
			Case "Undeliverable Email"
				Find_Contact_Form objContactForm 
				Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_INSD").Value = objContactForm.Form.GetFieldValue ("//data/ClaimPrimaryEmail")
				Task.GetAttributeObject("BPA_ADJUSTING_HOME_INSD").Value = objContactForm.Form.GetFieldValue ("//data/ClaimPrimaryHome")
				Task.GetAttributeObject("BPA_ADJUSTING_CELL_INSD").Value = objContactForm.Form.GetFieldValue ("//data/ClaimPrimaryCell")
				Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_ALT").Value = objContactForm.Form.GetFieldValue ("//data/ClaimSecondaryEmail")
				Task.GetAttributeObject("BPA_ADJUSTING_HOME_ALT").Value = objContactForm.Form.GetFieldValue ("//data/ClaimSecondaryHome")
				Task.GetAttributeObject("BPA_ADJUSTING_CELL_ALT").Value = objContactForm.Form.GetFieldValue ("//data/ClaimSecondaryCell")

'				Task.AssignedTo = AccountsLookup.GetUser("Black Pearl User")
				Task.Folder.UpdateType("Automated Correspondence")
				Document.UpdateType("UDEM")
					Select Case strDrawer
						Case "CPCL"
							Logger.Info "Hub - Indexed - Undeliverable Email - CPIC"
							Hub = 14
						Case "SFCL"
							Logger.Info "Hub - Indexed - Undeliverable Email - SFIC"
							Hub = 15
						Case Else
							Logger.Info "Hub - Indexed - Undeliverable Email - SFPC"
							Hub = 16
					End Select
			Case "Not BPA"
				Logger.Info "Hub - Not BPA"
				Hub = 12
				Exit Function
			Case "Junk Mail"
				Logger.Info "Hub - Junk Mail"
				Hub = 9
				Exit Function
			Case Else
				Logger.Info "Hub - Research"
				Hub = 11
				Exit Function
		End Select		
		strParameter = "[%MESSAGE_TO%]:" & Document.GetAttributeObject("ER_FROMADDR").Value
		Update_CDS_Status File.GetAttributeObject("ADJUSTMENT_STATUS").Value, strParameter
	Case "Correspondence Reviewed"
		Logger.Info "Hub - Correspondence Reviewed"
		Select Case strAdjust
			Case "Need More Information"
				Logger.Info "Hub - Need More Information"
															strInsd = Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_INSD").Value:										Logger.Info "Adjusting_Work - Insured Email - " & strInsd
				Dim strSender:					strSender = Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_RECEIVED_FROM").Value:			Logger.Info "Adjusting_Work - From - " & strSender
				Dim strSendTo:					strSendTo = Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_SEND_TO").Value:						Logger.Info "Adjusting_Work - Send To - " & strSendTo
				Dim strComments:			strComments = Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_COMMENT1").Value  & " " & Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_COMMENT2").Value:				Logger.Info "Adjusting_Work - Comments - " & strComments
				strComments = Replace(strComments, "'", "''")
				Select Case strDrawer
					Case "CPCL"
						strPolicyDrawer = "CPUW"
					Case "SFCL"
						strPolicyDrawer = "SFUW"
					Case Else
						strPolicyDrawer = "SPCU"
				End Select		
				Logger.Info "Hub - Policy Drawer - " & strPolicyDrawer
				strNote = "We recevied documentation, but still need the following additional information: <br/>" & strComments:						Logger.Info "Hub - Note - " & strNote
				File.GetAttributeObject("ADJUSTMENT_STATUS").Value = "Awaiting Further Information" 

				Select Case strSendTo
					Case "Insured"
						strTo = strInsd
					Case "Sender"
						strTo = strSender
					Case Else
						strTo = strInsd
						strCc = strSender
				End Select
				File.ClearFileMarkById(33462655)'BPA First Follow Up Sent 					Production 59519388			Development 33462655
				File.ClearFileMarkById(33462656)'BPA Second Follow Up Sent				Production 59519450 			Development 33462656
				File.ClearFileMarkById(33462657)'BPA Final Follow Up Sent					Production 59519507			Development 33462657

				strParameter = "[%MESSAGE_TO%]:" & strTo & "|[%MESSAGE_CC%]:" & strCc & "|[%MESSAGE_COMMENT%]:" & strComments 
				Hub = 9
			Case "Send to Claim Print Check Review"
				Logger.Info "Hub - Send to Claim Print Check Review"
				File.ClearFileMarkById(33462655)'BPA First Follow Up Sent 					Production 59519388			Development 33462655
				File.ClearFileMarkById(33462656)'BPA Second Follow Up Sent				Production 59519450 			Development 33462656
				File.ClearFileMarkById(33462657)'BPA Final Follow Up Sent					Production 59519507			Development 33462657
				strNote = "The claims package has been sent for management review.":						Logger.Info "Hub - Note - " & strNote
				Task.Description = "BPA Claim Package Review"
				File.GetAttributeObject("ADJUSTMENT_STATUS").Value = "Under Manager Review" 
				Dim strMoved:								strMoved = "Not Moved"
				Dim strPageMarkColor: 				strPageMarkColor = "34, 139, 34"
				Dim arrayFolders, arrayDocs, arrayPages, x, y, z
				arrayFolders = File.GetFolders
				Logger.Info "Adjusting_Work - Folders - UBound - " & UBound(arrayFolders) + 1
				For x = 0 To UBound(arrayFolders)
					Logger.Info "Adjusting_Work - Folder - " & (x + 1) & " - " & arrayFolders(x).Name
					arrayDocs = arrayFolders(x).GetDocuments
					Logger.Info "Adjusting_Work -Docs - UBound - " & UBound(arrayDocs) + 1
						For y = 0 To UBound(arrayDocs)
							Logger.Info "Adjusting_Work - Doc - " & (y + 1) & " - " & arrayDocs(y).Name
							arrayPages = arrayDocs(y).GetPages
							Logger.Info "Adjusting_Work - Page - UBound - " & UBound(arrayPages) + 1
								For z = 0 To UBound(arrayPages)
									Logger.Info "Adjusting_Work - Page - " & (z + 1) & " - " & arrayPages(z).PageId & " - Has Mark - " & arrayPages(z).HasPageMark(strPageMarkColor)
									If strMoved = "Not Moved" Then
										If arrayPages(z).HasPageMark(strPageMarkColor) = "True" Then
											Set objPage = arrayPages(z)
											Logger.Info "Adjusting_Work - Page Info - " & objPage.PageId
											IRSL.ChangeTaskTarget Task, objPage
'											Task.SetTarget objPage
											Logger.Info "Adjusting_Work - Task Moved"
											objPage.ClearPageMarkById(33462976)
											strMoved = "Moved"
											x = UBound(arrayFolders)
											y = UBound(arrayDocs)
											z = UBound(arrayPages)
										End If
									End If
								Next
						Next
					arrayDocs = Null
				Next
				Hub = 13
			Case Else
				Logger.Info "Hub - End of Flow"
				strNote = "The correspondence received has been release to End of Flow."
				Dim  intFileId : intFileId = Task.File.Id
				Dim strStepProgName, arrTasks, objTask, intTasks
				Logger.Info "Hub - Id for this file is [" & CStr(intFileId) & "]"
				arrFlows = Array("BPAW", "CLM-PRT")	:																			Logger.Info "Hub - Flow - " & arrFlows(0)
				For y = 0 To UBound(arrFlows)
					Dim arrSteps:											arrSteps = FlowLookup.FindSteps(arrFlows(y)):									Logger.Info "Hub - Steps - " & UBound(arrSteps) + 1 & " - " & " in flow " & arrFlows(y)
					For x = 0 To UBound(arrSteps)
						strStepProgName = arrSteps(x).ProgrammaticName				'found in the StepDef table                
						                
						On Error Resume Next
						Dim objStep : Set objStep = FlowLookup.FindStep(arrFlows(y), strStepProgName)
						Logger.Info "Hub - GetTasks - Create A Step Object? [" & IsObject(objStep) & "] for Step [" & objStep.Name & "] in Flow [" & objStep.Flow.Name & "]"
						                                
						If objStep Is Nothing Or Not IsObject(objStep) Then                       
							Logger.Info "Hub - GetTasks - Unable to Create Step Object for [" & CStr(strStepProgName) & "] In Flow [" & strFlowProgName & "]"
							Exit Function
						End If
              				'***Create Flow SearchCondition
							Dim searchCondition : Set searchCondition = objStep.Flow.CreateTaskSearchCondition()
							Logger.Info "Hub - GetTasks - Is Search Flow Condition an Object? [" & IsObject(searchCondition) & "]"

							'***Now Add the Step
							searchCondition.AddStep objStep.Flow.ProgrammaticName, objStep.ProgrammaticName
							searchCondition.SetIncludeDiary True
							                                
							On Error Goto 0                

							'***Error Adding the Step
							If Not Err.Number = 0 Then 
								Logger.Info "Get Task - Errored Occurred Adding Step [" & objStep.ProgrammaticName & "]"
							                
							'***Added Step Successfully
							Else

								'***Now Add the File
								If Not IsEmpty(intFileId) Then                    
									searchCondition.AddFile intFileId
									Logger.Info "Hub - GetTasks - Added File Filter [" & CStr(intFileId) & "]"
								End If                                    
								                                
								On Error Resume Next
								'***Now get the Tasks
								arrTasks = objStep.Flow.GetTasks(searchCondition)
								Logger.Info "Hub - GetTasks - Did we find Tasks? [" &  (UBound(arrTasks) + 1) & "]"
								On Error Goto 0
								 
								 intTasks = intTasks + (UBound(arrTasks)  + 1)              
								
							End If
					Next
				Next
				Logger.Info "Hub - Total Tasks - " & intTasks
				Set objStep = Nothing
				intFileId = ""
				strFlowProgName = ""
				strStepProgName = ""

				If intTasks = 0 Then
					File.GetAttributeObject("ADJUSTMENT_STATUS").Value = "Awaiting Further Information" 
				End If
				Hub = 9
		End Select
		AutoNote strDrawer, strFile, strFlow, strStep, strUser, strPriority, strTaskDescription, strNote
		Update_CDS_Status File.GetAttributeObject("ADJUSTMENT_STATUS").Value, strParameter
	Case "Package Declined"
		Logger.Info "Hub - Package Declined"
		Task.Description = "Package Declined - Please review."
		Task.GetAttributeObject("CHECK_AMOUNT").Value = "0.00"
			Select Case strDrawer
				Case "CPCL"
					Logger.Info "Hub - Package Declined - CPIC"
					Hub = 5
				Case "SFCL"
					Logger.Info "Hub - Package Declined - SFIC"
					Hub = 6
				Case Else
					Logger.Info "Hub - Package Declined - SFPC"
					Hub = 7
			End Select
	Case "Undeliverable Reviewed"
		Find_Contact_Form objContactForm 
		Dim strUndeliverable:						strUndeliverable = Task.GetAttributeObject("BPA_UNDELIVERABLE_EMAIL_TYPE").Value:							Logger.Info "Undeliverable_Email - Undeliverable Email - " & strUndeliverable
		Dim strUndeliverableContact:		strUndeliverableContact = Task.GetAttributeObject("BPA_UNDELIVERABLE_EMAIL_CONTACT").Value:	Logger.Info "Undeliverable_Email - Undeliverable Contact - " & strUndeliverableContact
		strEmail = Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_INSD").Value:																												Logger.Info "Undeliverable_Email - Email - " & strEmail
		strPhone = Task.GetAttributeObject("BPA_ADJUSTING_HOME_INSD").Value:																												Logger.Info "Undeliverable_Email - Phone - " & strPhone
		strCell = Task.GetAttributeObject("BPA_ADJUSTING_CELL_INSD").Value:																														Logger.info "Undeliverable_Email - Contact Cell - " & strCell
		strAltEmail = Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_ALT").Value:																											Logger.info "Undeliverable_Email - Alt Email - " & strAltEmail
		strAltHome = Task.GetAttributeObject("BPA_ADJUSTING_HOME_ALT").Value:																											Logger.info "Undeliverable_Email - Alt Home - " & strAltHome
		strAltCell = Task.GetAttributeObject("BPA_ADJUSTING_CELL_ALT").Value:																													Logger.info "Undeliverable_Email - Alt Cell - " & strAltCell
		strOtherEmail = Task.GetAttributeObject("BPA_UNDELIVERABLE_EMAIL_OTHER").Value:																						Logger.info "Undeliverable_Email - Other Email - " & strOtherEmail
		
		strOldEmail = objContactForm.Form.GetFieldValue ("//data/ClaimPrimaryEmail"):																									Logger.info "Undeliverable_Email - Original Contact Email - " & strOldEmail
		strOldPhone = objContactForm.Form.GetFieldValue ("//data/ClaimPrimaryHome"):																								Logger.info "Undeliverable_Email - Original Contact Phone - " & strOldPhone
		stOldCell = objContactForm.Form.GetFieldValue ("//data/ClaimPrimaryCell"):																											Logger.info "Undeliverable_Email - Contact Cell - " & stOldCell
		strOldAltEmail = objContactForm.Form.GetFieldValue ("//data/ClaimSecondaryEmail"):																							Logger.info "Undeliverable_Email - Alt Email - " & strOldAltEmail
		strOldAltHome = objContactForm.Form.GetFieldValue ("//data/ClaimSecondaryHome"):																						Logger.info "Undeliverable_Email - Alt Home - " & strOldAltHome
		strOldAltCell = objContactForm.Form.GetFieldValue ("//data/ClaimSecondaryCell"):																								Logger.info "Undeliverable_Email - Alt Cell - " & strOldAltCell

		Dim strDate:							strDate = Task.Document.DocumentDate:																															Logger.Info "Undeliverable_Email - Date - " & strDate
		Dim strParameters:				strParameters = ""
		Dim strResendTo:						strResendTo = ""
		strChanges = ""
		strAltName = ""
		strUser = Task.FromUser.Name:																																														Logger.Info "Undeliverable_Email - From User - " & strUser
		'strUser = "rschwinn":																																																			Logger.Info "Undeliverable_Email - From User - " & strUser

		Select Case strUndeliverable
			Case "Email Address Updated"
				If strEmail <> strOldEmail Then
					Logger.Info "Undeliverable_Email - Email Correction"
					strChanges = strChanges & "<p> <b>Original Insured Email Address:</b> " & strOldEmail & "<br/> <b>Revised Insured Email Address:</b> " & strEmail  & "</p>"
					If strEmail = "" Then
						strEmail = strOldEmail
					End If
					strResendTo = strResendTo & ";" & strEmail
					objContactForm.Form.SetFieldValue "//data/ClaimPrimaryEmail", strEmail
				End If
				If strPhone <> strOldPhone Then
					Logger.Info "Undeliverable_Email - Phone Correction"
					strChanges = strChanges & "<p> <b>Original Insured Home Phone:</b> " & strOldPhone & "<br/> <b>Revised Insured Home Phone:</b> " & strPhone  & "</p>"
					If strPhone = "" Then
						strPhone = strOldPhone
					End If
					objContactForm.Form.SetFieldValue "//data/ClaimPrimaryHome", strPhone
				End If
				If strCell <> stOldCell Then
					Logger.Info "Undeliverable_Email - Cell Correction"
					strChanges = strChanges & "<p> <b>Original Insured Cell Phone:</b> " & stOldCell & "<br/> <b>Revised Insured Cell Phone:</b> " & strCell  & "</p>"
					If strCell = "" Then
						strCell = stOldCell
					End If
					objContactForm.Form.SetFieldValue "//data/ClaimPrimaryCell", strCell
				End If
				If strAltEmail <> strOldAltEmail Then
					Logger.Info "Undeliverable_Email - Alt Email Correction"
					strChanges = strChanges & "<p> <b>Original Alternate Contact Email:</b> " & strOldAltEmail & "<br/> <b>Revised Alternate Contact Email:</b> " & strAltEmail  & "</p>"
					If strAltEmail = "" Then
						strAltEmail = strOldAltEmail
					End If
					objContactForm.Form.SetFieldValue "//data/ClaimSecondaryEmail", strAltEmail
					strResendTo = strResendTo & ";" & strAltEmail
				End If
				If strAltHome <> strOldAltHome Then
					Logger.Info "Undeliverable_Email - Alt Home Phone Correction"
					strChanges = strChanges & "<p> <b>Original Alternate Contact Home Phone:</b> " & strOldAltHome & "<br/> <b>Revised Alternate Contact Home Phone:</b> " & strAltHome  & "</p>"
					If strAltHome = "" Then
						strAltHome = strOldAltHome
					End If
					objContactForm.Form.SetFieldValue "//data/ClaimSecondaryHome", strAltHome
				End If
				If strAltCell <> strOldAltCell Then
					Logger.Info "Undeliverable_Email - Alt Cell Phone Correction"
					strChanges = strChanges & "<p> <b>Original Alternate Contact Cell Phone:</b> " & strOldAltCell & "<br/> <b>Revised Alternate Contact Cell Phone:</b> " & strAltCell  & "</p>"
					If strAltCell = "" Then
						strAltCell = strOldAltCell
					End If
					objContactForm.Form.SetFieldValue "//data/ClaimSecondaryCell", strAltCell
				End If
				If strOtherEmail <> "" Then
					Logger.Info "Undeliverable_Email - Other Email Correction"
					strResendTo = strResendTo & ";" & strOtherEmail
				End If
				
				If strChanges <> "" Then
					Logger.Info "Undeliverable_Email - Correction Start"
					Select Case File.Drawer.Name
						Case "CPCL"
							strNoteDrawer = "CPUW"
						Case "SFCL"
							strNoteDrawer = "SFUW"
						Case Else
							strNoteDrawer = "SPCU"
					End Select	
					
					Update_CDS_Contact Task.File.FullFileNumber, strEmail, strPhone, strCell
					Update_CDS_Alternate_Contact Task.File.FullFileNumber, strAltEmail, strAltHome, strAltCell, strAltName

					Logger.Info "Undeliverable_Email - Note Drawer - " & strNoteDrawer
					strPolicy = objContactForm.Form.GetFieldValue("//data/PolicyNumber"):				Logger.Info "Undeliverable_Email - File - " & strPolicy
					strFlow = "INFORCE":																											Logger.Info "Undeliverable_Email - File - " & strFlow
					strStep = "S_b66c62a0-d455-4e11-93c3-15afb337309f":											Logger.Info "Undeliverable_Email - File - " & strStep
					strNote = "<p>During a claim call the insured indicated that our records are incorrect and need to be updated with the information below.</p>" & strChanges & _
											"<p>Please review this information and make the appropriate changes. Contact the insured or agent if additional information is needed to make the changes."
					Logger.Info "Undeliverable_Email - Changes - " & strNote
					AutoNote strNoteDrawer, strPolicy, strFlow, strStep, strUser,"4", "BPA Corrections", strNote
					strFlow = ""
					strStep = ""
					strNote = "<p>An undeliverable email notification was received. Alternate contact information was found and we have updated the contact information in CDS and sent a task for CSR to review for possible changes on the policy..</p>" & strChanges
					AutoNote File.Drawer.Name, Task.File.FullFileNumber, strFlow, strStep, strUser,"", "", strNote
					objContactForm.Form.Save
					strStatus = "Undeliverable - Contact Updated"
				End If
				If strResendTo <> "" Then
					strParameters = "[%MESSAGE_TO%]:" & strResendTo & "[%RESEND_FROM%]:" & strDate
				End If
				Update_CDS_Status strStatus, strParameters
				Hub = 9
			Case "Attempted to contact insured"
				File.GetAttributeObject("ADJUSTMENT_STATUS").Value = "Undeliverable - Attempted Contact"
				Task.GetAttributeObject("BPA_UNDELIVERABLE_EMAIL_TYPE").Value = ""
				Task.DateAvailable = DateAdd("h",4,DateValue(Date()+1)):				Logger.Info "Undeliverable_Email - Uncontacted - Date Available - " & Task.DateAvailable
				Task.Description = "Another attempt to get alternate contact info."
				Select Case strDrawer
					Case "CPCL"
						Logger.Info "Undeliverable_Email - Another contact attempt - CPIC"
						Hub = 14
					Case "SFCL"
						Logger.Info "Undeliverable_Email - Another contact attempt - SFIC"
						Hub = 15
					Case Else
						Logger.Info "Undeliverable_Email - Another contact attempt - SFPC"
						Hub = 16
				End Select
				strNote = "Unable to contact insured to procure alternate contact information. Task has been rescheduled for tomorrow."
				AutoNote File.Drawer.Name, Task.File.FullFileNumber, strFlow, strStep, strUser,"", "", strNote
				strStatus = "Undeliverable - Attempted Contact"
				Update_CDS_Status strStatus, strParameters

			End Select
		Logger.Info "Undeliverable_Email - End"
		
	Case Else
		Select Case strDrawer
			Case "CPCL"
				Logger.Info "Hub - Else - CPIC"
				Hub = 5
			Case "SFCL"
				Logger.Info "Hub - Else - SFIC"
				Hub = 6
			Case Else
				Logger.Info "Hub - Else - SFPC"
				Hub = 7
		End Select
End Select 
Logger.Info "Hub - End"
End Function

'******************** Black Pearl Workflow - Development *************************
'******************** Hub ***************************************************
'******************** Email_Indexing ******************************************
Function Email_Indexing()
Logger.Info "Email_Indexing - Start"

Dim strSQL, strDB, sResult
Dim strFile:		strFile = Task.File.FullFileNumber

If Left(strFile,3) = "~ER" Then
	Logger.Info "Email_Indexing - Temp File"
Else
	Logger.Info "Email_Indexing - Email Reply"
	strSQL = "select Utility.dbo.[Lookup_Claim_Company] ('" + strFile + "')":																Logger.Info "Email_Indexing - SQL - " & strSQL
	strDB = "PROVIDER=SQLOLEDB;Data Source=ir5dbdev;Initial Catalog=Utility;uid=sa;PWD=cpic0742":			Logger.Info "Email_Indexing - DB - " & strDB
	Call_Database strSQL, strDB, sResult
	Index_File strFile, sResult
End If
Email_Indexing = 1

Logger.Info "Email_Indexing - End"
End Function

'******************** Black Pearl Workflow - Development *************************
'******************** Hub ***************************************************
'******************** Adjusting Work ******************************************
Function  Adjusting_Work()
Logger.Info "Adjusting_Work - Start"

Dim strType:						strType = Task.GetAttributeObject("BPA_ADJUSTING_TYPE").Value:													Logger.Info "Adjusting_Work - Type - " & strType
Dim strInsd:						strInsd = Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_INSD").Value:										Logger.Info "Adjusting_Work - Insured Email - " & strInsd
Dim strSender:					strSender = Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_RECEIVED_FROM").Value:			Logger.Info "Adjusting_Work - From - " & strSender
Dim strSendTo:					strSendTo = Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_SEND_TO").Value:						Logger.Info "Adjusting_Work - Send To - " & strSendTo
Dim strComments:			strComments = Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_COMMENT1").Value  & " " & Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_COMMENT2").Value:				Logger.Info "Adjusting_Work - Comments - " & strComments
Dim strErrrors:					strErrors = ""
Dim strDrawer:					strDrawer = Task.File.Drawer.Name:																											Logger.Info "Adjusting_Work - Drawer - " & strDrawer
Dim strPaymentType:		strPaymentType = Task.GetAttributeObject("TYPE_OF_CHECK_REQUEST").Value:							Logger.Info "Adjusting_Work - Payment Type - " & strPaymentType
Dim strPayment:				strPayment = Task.GetAttributeObject("CHECK_AMOUNT").Value:													Logger.Info "Adjusting_Work - Payment Amount - " & strPayment
Dim strPolicyDrawer, strFile, strFlow, strStep, strUser, strPriority, strTaskDescription, strNote, strSubject, strFrom, strTo, strCc, strBcc, strMessage, strImage, strCompany, strCompanyName
Task.GetAttributeObject("ADJUSTMENT_STATUS").Value = "Correspondence Reviewed" 
Task.GetAttributeObject("DELIVERY_METHOD").Value = "Email":																								Logger.Info "Adjusting_Work - Delivery Method - " & Task.GetAttributeObject("DELIVERY_METHOD").Value

If strType = "" Then
	strErrors = strErrors & VbCr & "Please select Set Task Attributes and complete the first field. "
End If
If strType = "Need More Information" Then
	If(strSendTo = "" Or strSendTo = "Neither") Then
		strErrors = strErrors & VbCr & "Please select who you want to send this email to. "
	End If
	If (strSendTo = "Insured" Or strSendTo = "Both") And (strInsd = "No Email on Record" Or strInsd = "") Then
		strErrors = strErrors & VbCr & "You are trying to send this email to the insured, but there is no email address for them. Please enter the insureds email address or change your email options."
	End If
	If (strSendTo = "Sender" Or strSendTo = "Both") And (strSender = "No Email on Record" Or strSender = "") Then
		strErrors = strErrors & VbCr & "You are trying to send this email to the person who sent it, but there is no email address for them. Please enter the senders email address or change your email options."
	End If
	If (strSendTo = "" Or strSendTo = "Neither") And strComments <> "" Then
		strErrors = strErrors & VbCr & "You entered a comment, but did not select that you want to send an email."
	End If
	If strComments = "" Then
		strErrors = strErrors & VbCr & "Please add comments to let the recipient know what you still need."
	End If
	If strPaymentType <> "" Then
		strErrors = strErrors & VbCr & "You selected 'Needs more information', but selected a payment type. Please correct one of these options. "
	End If 
	If strDelivery <> "" Then
		strErrors = strErrors & VbCr & "You selected 'Needs more information', but selected a delivery method. Please correct one of these options. "
	End If
Else
	If strComments <> " " Then
		strErrors = strErrors & VbCr & "You selected 'Send to Claim Print Check Review', but entered comments. Please remove the comments."
	End If
	If strPayment = "0" And strPaymentType <> "No Payment" Then
		strErrors = strErrors & VbCr & "Please enter the amount of the payment. " & strPayment
	End If
	If strPayment > "0.00" And strPaymentType = "No Payment" Then
		strErrors = strErrors & VbCr & "Please change the payment amount to $0 or select the correct payment type."
	End If
	If strPaymentType = "" Then
		strErrors = strErrors & VbCr & "Please select the type of payment package being submitted."
	End If
	If strType = "Send to Claim Print Check Review" Then
		'Find Pages Marked with specific colors 
		Dim strMoved:								strMoved = "Not Found"
		Dim strPageMarkColor: 				strPageMarkColor = "34, 139, 34"
		Dim arrayFolders, arrayDocs, arrayPages, x, y, z, strPageDoc
		arrayFolders = File.GetFolders
		Logger.Info "Adjusting_Work - Folders - UBound - " & UBound(arrayFolders) + 1
		For x = 0 To UBound(arrayFolders)
			Logger.Info "Adjusting_Work - Folder - " & (x + 1) & " - " & arrayFolders(x).Name
			arrayDocs = arrayFolders(x).GetDocuments
			Logger.Info "Adjusting_Work -Docs - UBound - " & UBound(arrayDocs) + 1
				For y = 0 To UBound(arrayDocs)
					Logger.Info "Adjusting_Work - Doc - " & (y + 1) & " - " & arrayDocs(y).Name
					arrayPages = arrayDocs(y).GetPages
					Logger.Info "Adjusting_Work - Page - UBound - " & UBound(arrayPages) + 1
						For z = 0 To UBound(arrayPages)
							Logger.Info "Adjusting_Work - Page - " & (z + 1) & " - " & arrayPages(z).PageId & " - Has Mark - " & arrayPages(z).HasPageMark(strPageMarkColor)
							If strMoved = "Not Found" Then
								If arrayPages(z).HasPageMark(strPageMarkColor) = "True" Then
									Set objPage = arrayPages(z)
									Logger.Info "Adjusting_Work - Page Info - " & objPage.PageId & " - " & objPage.Document.Name
									strPageDoc = objPage.Document.Name
									If strPageDoc = "1PCR" Or strPageDoc = "CLDL" Or strPageDoc = "NPLT" Or strPageDoc = "ROR" Or strPageDoc = "MISC" Or strPageDoc = "CHKS" Or strPageDoc = "CREM" Then
										Logger.Info "Adjusting_Work - Correct DocType"
									Else
										Logger.Info "Adjusting_Work - Incorrect DocType"
										strErrors = strErrors & VbCr & "The 'Send BPA Claim Package to Check Print Review' page mark is not on a document type allowed in that flow. Please change the document type or the location of the page mark."
									End If
									strMoved = "Found"
									x = UBound(arrayFolders)
									y = UBound(arrayDocs)
									z = UBound(arrayPages)
								End If
							End If
						Next
				Next
			arrayDocs = Null
		Next
		If strMoved = "Not Found" Then
			Logger.Info "Adjusting_Work - No Page Mark"
			strErrors = strErrors & VbCr & "A 'Send BPA Claim Package to Check Print Review' page mark was not found in the file. Please set the page mark on the page you want to send for review."
		End If
	End If
End If

If strType = "End of Flow" Then
	Logger.Info "Adjusting_Work - End of Flow"
	Adjusting_Work = 1
	Exit Function
End If

If strErrors <> "" Then
	Logger.Info "Adjusting_Work - Errors - " & strErrors
	Adjusting_Work = strErrors
	Exit Function
End If

Adjusting_Work = 1
Logger.Info "Adjusting_Work - End"
End Function

'******************** Black Pearl Workflow - Development *************************
'******************** Hub ***************************************************
'******************** New Loss ***********************************************
Function New_Loss()
Logger.Info "New_Loss - Start"

Select Case Task.GetAttributeObject("BPA_CONTACT").Value
	Case "" 
		New_Loss = "Please use Set Task Attributes to send this to the correct place."
		Exit Function
	Case Else
		Task.GetAttributeObject("ADJUSTMENT_STATUS").Value = "Contact"
		New_Loss = 1
End Select

Logger.Info "New_Loss - End"
End Function

'******************** Black Pearl Workflow - Development *************************
'******************** Hub ***************************************************
'******************** Indexing ***********************************************
Function Indexing()
Logger.Info "Indexing - Start"
Dim strIndexType:			strIndexType = Task.GetAttributeObject("BPA_INDEXING_TYPE").Value:									Logger.Info "Indexing - Type - " & strIndexType
Task.GetAttributeObject("ADJUSTMENT_STATUS").Value = "Indexed"

Select Case strIndexType
	Case "Adjusting Work"
		Logger.Info "Indexing - Adjusting Work"
		Task.Description = "Adjusting Work"
		File.GetAttributeObject("ADJUSTMENT_STATUS").Value = "Correspondence Awaiting Review"
	Case "Research"
		Logger.Info "Indexing - Research"
		Task.Description = "Research"
		File.GetAttributeObject("ADJUSTMENT_STATUS").Value = "Research"
	Case "Undeliverable Email"
		Logger.Info "Indexing - Undeliverable Email"
		Task.Description = "Undeliverable Email"
		File.GetAttributeObject("ADJUSTMENT_STATUS").Value = "Undeliverable Email"
	Case "Not BPA"
		Logger.Info "Indexing - Not BPA"
		Task.Description = "Not BPA"
	Case Else
		Logger.Info "Indexing - Junk Mail"
		Task.Description = "Junk Mail"
End Select
Indexing = 1
Logger.Info "Indexing - End"
End Function

'******************** Black Pearl Workflow - Development *************************
'******************** Hub ***************************************************
'******************** Research ***********************************************
Function Research()
Logger.Info "Research - Start"

Select Case Task.GetAttributeObject("BPA_RESEARCH").Value 
	Case "Adjusting Work" 
		Logger.Info "Research - Adjusting Work"
		Task.GetAttributeObject("ADJUSTMENT_STATUS").Value = "Indexed"
		Task.GetAttributeObject("BPA_INDEXING_TYPE").Value = "Adjusting Work"
		File.GetAttributeObject("ADJUSTMENT_STATUS").Value = "Correspondence Awaiting Review"
		Task.Description = "Adjusting Work"
	Case "Not BPA"
		Logger.Info "Research - Not BPA"
		Task.GetAttributeObject("ADJUSTMENT_STATUS").Value = "Indexed"
		Task.GetAttributeObject("BPA_INDEXING_TYPE").Value = "Not BPA"
		Task.Description = "Not BPA"
	Case Else
		Logger.Info "Research - Junk Mail"
		Task.GetAttributeObject("ADJUSTMENT_STATUS").Value = "Indexed"
		Task.GetAttributeObject("BPA_INDEXING_TYPE").Value = "Junk Mail"
		Task.Description = "Junk Mail"
	End Select
Research = 1
Logger.Info "Research - End"
End Function

'******************** Black Pearl Workflow - Development *************************
'******************** Hub ***************************************************
'******************** Undeliverable Email **************************************
Function Undeliverable_Email()
Logger.Info "Undeliverable_Email - Start"

Dim strUndeliverable:			strUndeliverable = Task.GetAttributeObject("BPA_UNDELIVERABLE_EMAIL_TYPE").Value:									Logger.Info "Undeliverable_Email - Undeliverable Email - " & strUndeliverable
Dim strEmail:							strEmail = Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_INSD").Value:															Logger.Info "Undeliverable_Email - Email - " & strEmail
Dim strPhone:						strPhone = Task.GetAttributeObject("BPA_ADJUSTING_HOME_INSD").Value:															Logger.Info "Undeliverable_Email - Phone - " & strPhone
Dim strContact:					strContact = Task.GetAttributeObject("BPA_UNDELIVERABLE_EMAIL_CONTACT").Value:										Logger.Info "Undeliverable_Email - Phone - " & strContact
Dim strErrors:						strErrors = ""
Logger.Info "Undeliverable_Email - Email - " & File.GetAttributeObject("BPA_ADJUSTING_EMAIL_INSD").Value
Logger.Info "Undeliverable_Email - Phone - " & File.GetAttributeObject("BPA_ADJUSTING_HOME_INSD").Value

If strContact = "" Then
	strErrors = strErrors & VbCr & "Please select who the undeliverable email was sent to."
End If
If strUndeliverable = "Email Address Updated" And strEmail = File.GetAttributeObject("BPA_ADJUSTING_EMAIL_INSD").Value Then
	strErrors = strErrors & VbCr & "You did not correct the email address. Please make appropriate corrections."
End If
If strEmail = "" Then
	strErrors = strErrors & VbCr & "The insureds email address cannot be blank."
End If
If strErrors <> "" Then
	Undeliverable_Email = strErrors
	Exit Function
End If

Task.GetAttributeObject("ADJUSTMENT_STATUS").Value = "Undeliverable Reviewed"
Undeliverable_Email = 1
Logger.Info "Undeliverable_Email - End"
End Function

'******************** Black Pearl Workflow - Development *************************
'******************** Hub ***************************************************
'******************** Find Contact Form ****************************************
Function Find_Contact_Form(objContactForm)
Logger.Info "Find_Contact_Form - Start"

Dim arrFolders, arrDocs, arrPages, strMortgage
arrFolders = Task.File.GetFoldersOfType("1st Party Claim Folder",False)
Set objFolder = arrFolders(0)
arrDocs = objFolder.GetDocumentsOfType("CIO",False)
Set objDoc = arrDocs(0)
arrPages = objDoc.GetPages()
Set objContactForm = arrPages(0)

Logger.Info "Find_Contact_Form - Folder - " & objFolder.Id
Logger.Info "Find_Contact_Form - Doc - " & objDoc.Id
Logger.Info "Find_Contact_Form - Page - " & objContactForm.PageId

Logger.Info "Find_Contact_Form - End"
End Function

'******************** Black Pearl Workflow - Development *************************
'******************** Hub ***************************************************
'******************** Get Contact Form Info **************************************
Function Get_Contact_Form_Info()
Logger.Info "Get_Contact_Form_Info - Start"

Find_Contact_Form objContactForm

Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_INSD").Value = objContactForm.Form.GetFieldValue ("//data/ClaimPrimaryEmail"):					Logger.Info "Get_Contact_Form_Info - Insured Email - " & Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_INSD").Value
Task.GetAttributeObject("BPA_ADJUSTING_HOME_INSD").Value = objContactForm.Form.GetFieldValue ("//data/ClaimPrimaryHome"):				Logger.Info "Get_Contact_Form_Info - Insured Home - " & Task.GetAttributeObject("BPA_ADJUSTING_HOME_INSD").Value
Task.GetAttributeObject("BPA_ADJUSTING_CELL_INSD").Value = objContactForm.Form.GetFieldValue ("//data/ClaimPrimaryCell"):						Logger.Info "Get_Contact_Form_Info - Insured Cell - " & Task.GetAttributeObject("BPA_ADJUSTING_CELL_INSD").Value
Task.GetAttributeObject("BPA_ADJUSTING_NAME_ALT").Value = objContactForm.Form.GetFieldValue ("//data/ClaimSecondaryName"):				Logger.Info "Get_Contact_Form_Info - Alt Name - " & Task.GetAttributeObject("BPA_ADJUSTING_NAME_ALT").Value
Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_ALT").Value = objContactForm.Form.GetFieldValue ("//data/ClaimSecondaryEmail"):				Logger.Info "Get_Contact_Form_Info - Alt Email - " & Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_ALT").Value
Task.GetAttributeObject("BPA_ADJUSTING_HOME_ALT").Value = objContactForm.Form.GetFieldValue ("//data/ClaimSecondaryHome"):				Logger.Info "Get_Contact_Form_Info - Alt Home - " & Task.GetAttributeObject("BPA_ADJUSTING_HOME_ALT").Value
Task.GetAttributeObject("BPA_ADJUSTING_CELL_ALT").Value = objContactForm.Form.GetFieldValue ("//data/ClaimSecondaryCell"):						Logger.Info "Get_Contact_Form_Info - Alt Cell - " & Task.GetAttributeObject("BPA_ADJUSTING_CELL_ALT").Value
strMortgage = objContactForm.Form.GetFieldValue ("//data/Mortgage1Name1") & " " & objContactForm.Form.GetFieldValue ("//data/Mortgage1Name2") & " " & objContactForm.Form.GetFieldValue ("//data/Mortgage1Name3") & ", " & objContactForm.Form.GetFieldValue ("//data/Mortgage1Address") & ", " & objContactForm.Form.GetFieldValue ("//data/Mortgage1CityStateZip") & ", Loan # " & objContactForm.Form.GetFieldValue ("//data/Mortgage1Loan")
Task.GetAttributeObject("BPA_MORTGAGE_CORRECTION").Value = strMortgage:																							Logger.Info "Get_Contact_Form_Info - Name - " & Task.GetAttributeObject("BPA_MORTGAGE_CORRECTION").Value

If objContactForm.Form.GetFieldValue ("//data/CAT")	 <> "Not a CAT" And File.GetAttributeObject("HURRICANE").Value = "" Then
	File.GetAttributeObject("HURRICANE").Value = objContactForm.Form.GetFieldValue ("//data/CAT"):													Logger.Info "Get_Contact_Info - CAT - " & File.GetAttributeObject("HURRICANE").Value
End If


Logger.Info "Get_Contact_Form_Info - End"
End Function
'******************** Black Pearl Workflow - Development *************************
'******************** Hub ***************************************************
'******************** Get Contact Info **************************************
Function Get_Contact_Info()
Logger.Info "Get_Contact_Info - Start"
Dim  sResult, strArray, x

Logger.Info "Get_Contact_Info - File Number - " & Task.File.FullFileNumber
Dim strSQL:			strSQL = "exec Utility.dbo.CDS_Claim_All_Contact_Info '" + File.FullFileNumber + "'":											Logger.Info "Get_Contact_Info - SQL - " & strSQL
Dim strDB:				strDB = "PROVIDER=SQLOLEDB;Data Source=ir5dbdev;Initial Catalog=Utility;uid=sa;PWD=cpic0742":				Logger.Info "Get_Contact_Info - DB - " & strDB

Call_Database strSQL, strDB, sResult

Logger.Info "Get_Contact_Info - " & sResult

	strArray = Split(sResult,";",-1,1) 
	For Each x in strArray
		Logger.Info "Get_Contact_Info - strArray - " & x
	Next
File.GetAttributeObject("POLICY_NUMBER").Value = strArray(0):								Logger.Info "Get_Contact_Info - Policy Number - " & File.GetAttributeObject("POLICY_NUMBER").Value
File.GetAttributeObject("CLAIM_DOL").Value = strArray(1):										Logger.Info "Get_Contact_Info - DOL - " & File.GetAttributeObject("CLAIM_DOL").Value
Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_INSD").Value = strArray(3):		Logger.Info "Get_Contact_Info - Insd Email - " & Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_INSD").Value
Task.GetAttributeObject("BPA_ADJUSTING_HOME_INSD").Value = strArray(4):		Logger.Info "Get_Contact_Info - Insd Home - " & Task.GetAttributeObject("BPA_ADJUSTING_HOME_INSD").Value
Task.GetAttributeObject("BPA_ADJUSTING_CELL_INSD").Value = strArray(5):			Logger.Info "Get_Contact_Info - Insd Cell - " & Task.GetAttributeObject("BPA_ADJUSTING_CELL_INSD").Value
Task.GetAttributeObject("BPA_MORTGAGE_CORRECTION").Value = strArray(6):	Logger.Info "Get_Contact_Info - Mortgage - " & Task.GetAttributeObject("BPA_MORTGAGE_CORRECTION").Value
Task.GetAttributeObject("BPA_ADJUSTING_NAME_ALT").Value = strArray(7):		Logger.Info "Get_Contact_Info - Alt Name - " & File.GetAttributeObject("BPA_ADJUSTING_NAME_ALT").Value
Task.GetAttributeObject("BPA_ADJUSTING_EMAIL_ALT").Value = strArray(8):		Logger.Info "Get_Contact_Info - Alt Email - " & File.GetAttributeObject("BPA_ADJUSTING_EMAIL_ALT").Value
Task.GetAttributeObject("BPA_ADJUSTING_HOME_ALT").Value = strArray(9):		Logger.Info "Get_Contact_Info - Alt Home - " & File.GetAttributeObject("BPA_ADJUSTING_HOME_ALT").Value
Task.GetAttributeObject("BPA_ADJUSTING_CELL_ALT").Value = strArray(10):			Logger.Info "Get_Contact_Info - Alt Cell - " & File.GetAttributeObject("BPA_ADJUSTING_CELL_ALT").Value

File.GetAttributeObject("BPA_ADJUSTING_EMAIL_INSD").Value = strArray(1):		Logger.Info "Get_Contact_Info - Email Insd - " & File.GetAttributeObject("BPA_ADJUSTING_EMAIL_INSD").Value
File.GetAttributeObject("BPA_ADJUSTING_HOME_INSD").Value = strArray(0):		Logger.Info "Get_Contact_Info - Cell Number - " & File.GetAttributeObject("BPA_ADJUSTING_HOME_INSD").Value
File.GetAttributeObject("BPA_MORTGAGE_CORRECTION").Value = strArray(5):		Logger.Info "Get_Contact_Info - Mortgage - " & File.GetAttributeObject("BPA_MORTGAGE_CORRECTION").Value
File.GetAttributeObject("BPA_ADJUSTING_EMAIL_ALT").Value = strArray(6):			Logger.Info "Get_Contact_Info - Alt Email - " & File.GetAttributeObject("BPA_ADJUSTING_EMAIL_ALT").Value
File.GetAttributeObject("BPA_PHONE_ALT").Value = strArray(7):								Logger.Info "Get_Contact_Info - Alt Phone - " & File.GetAttributeObject("BPA_PHONE_ALT").Value

If strArray(2) <> "Not a CAT" And File.GetAttributeObject("HURRICANE").Value = "" Then
	File.GetAttributeObject("HURRICANE").Value = strArray(2):									Logger.Info "Get_Contact_Info - CAT - " & Task.GetAttributeObject("HURRICANE").Value
End If
Logger.Info "Get_Contact_Info - End"
End Function

'******************** Black Pearl Workflow - Development *************************
'******************** Hub ***************************************************
'******************** Update CDS Contact **************************************
Function Update_CDS_Contact(strFile, strEmail, strCell, strPhone)
Logger.Info "Update_CDS_Contact - Start"

If strEmail = "Not on Record" Then
	strEmail = ""
End If
If strPhone = "Not on Record" Then
	strPhone = ""
End If
If strCell = "Not on Record" Then
	strCell = ""
End If
Dim strSQL:			strSQL = "exec dbo.IR_UpdateInsuredContact '" + strFile + "','" + strEmail + "','" + strCell + "','" + strPhone + "'":										Logger.Info "Update_CDS_Contact - SQL - "& strSQL
Dim strDB:				strDB = "PROVIDER=SQLOLEDB;Data Source=CDStest;Initial Catalog=Claim_Distribution_System;uid=sa;PWD=cpic0742":						Logger.Info "Update_CDS_Contact - DB - " & strDB

Call_Database strSQL, strDB, sResult

Logger.Info "Update_CDS_Contact - End"
End Function

'******************** Black Pearl Workflow - Development *************************
'******************** Hub ***************************************************
'******************** Update CDS Alternate Contact ******************************
Function Update_CDS_Alternate_Contact(strFile, strAltEmail, strAltHome, strAltCell, strAltName)
Logger.Info "Update_CDS_Alternate_Contact - Start"

If strAltEmail = "Not on Record" Then
	strAltEmail = ""
End If
If strAltHome = "Not on Record" Then
	strAltHome = ""
End If
If strAltCell = "Not on Record" Then
	strAltCell = ""
End If
If strAltName = "Not on Record" Then
	strAltName = ""
End If

Dim strSQL:			strSQL = "exec dbo.IR_UpdateAlternateContact '" + strFile + "','" + strAltEmail + "','" + strAltHome + "','" + strAltCell + "','" + strAltName + "'":		Logger.Info "Update_CDS_Alternate_Contact - SQL - "& strSQL
Dim strDB:				strDB = "PROVIDER=SQLOLEDB;Data Source=CDStest;Initial Catalog=Claim_Distribution_System;uid=sa;PWD=cpic0742":											Logger.Info "Update_CDS_Contact - DB - " & strDB

Call_Database strSQL, strDB, sResult

Logger.Info "Update_CDS_Alternate_Contact - End"
End Function

'******************** Black Pearl Workflow - Development *************************
'******************** Hub ***************************************************
'******************** Update CDS Status ***************************************
Function Update_CDS_Status(strStatus, strParameters)
Logger.Info "Update_CDS_Status - Start"
Logger.Info "Update_CDS_Status - Status - " & strStatus
Logger.Info "Update_CDS_Status - Parameters - " & strParameters

Dim strFile:			strFile = Task.File.FullFileNumber:																																									Logger.Info "Update_CDS_Status - File - " & strFile		
Dim strSQL:			strSQL = "exec BPA.updateClaimAdjustingStatus '" + strFile + "','" + strStatus + "','" + strParameters + "','ImageRight'":			Logger.Info "Update_CDS_Status - SQL - "& strSQL
Dim strDB:				strDB = "PROVIDER=SQLOLEDB;Data Source=cdstest;Initial Catalog=Claim_Distribution_System;uid=sa;PWD=cpic0742":		Logger.Info "Update_CDS_Status - DB - " & strDB

Call_Database strSQL, strDB, sResult
If sResult = "Bad Connection" Then
	sResult = "0"
Else 
	sResult = "1"
End If

Logger.Info "Update_CDS_Status - sResult - " & sResult
strSQL = "exec dbo.CDS_Status_Update '" + strFile + "','" + strStatus + "','" + strParameters + "','" + sResult + "'":							Logger.Info "Update_CDS_Status - SQL - " & strSQL
strDB = "PROVIDER=SQLOLEDB;Data Source=ir5dbdev;Initial Catalog=Utility;uid=sa;PWD=cpic0742":												Logger.Info "Update_CDS_Status - DB - " & strDB

Call_Database strSQL, strDB, sResult

Logger.Info "Update_CDS_Status - Result - " & sResult
Logger.Info "Update_CDS_Status - End"
End Function

'******************** Black Pearl Workflow - Development *************************
'******************** Hub ***************************************************
'******************** New Loss Contact Email ***********************************
Function New_Loss_Contact_Email(strEmail)
Dim strCompanyName, strImage, strCompany, strClaim, strCAT, strSubject, strEmailGreeting
strClaim = Task.File.FullFileNumber
strCAT = File.GetAttributeObject("HURRICANE").Value:	Logger.Info "New_Loss_Contact_Email - CAT - " & strCat
If strCAT = "" Then
	strSubject = "Claim"
Else
	strSubject = strCAT & " Claim"
End If

strImage = "<a href=http://www.blackpearladjusting.com><img src=http://www.capitol-preferred.com/images/260x120_CPIC-logo.gif alt='Black Pearl Adjusting. Click for home page.'></a>" 
strCompany = "<p>Black Pearl Adjusting<br/>Phone: 833.303.9764<br/><a href='www.blackpearladjusting.com'>www.blackpearladjusting.com</a></p>"

If Task.GetAttributeObject("BPA_CONTACT").Value = "Spoke with insured" Then
	strEmailGreeting = "<p>Claim Number:  " & strClaim & "</p><p>Thank you for your time today and choosing us to assist in adjusting your claim through the Fast-Track process. " & _
									  "Below is the information and steps that we discussed in submitting documentation directly to our office for review. " & _
									  "Please feel free to reac out to us at any time with questions.</p>"
Else
	strEmailGreeting = "<p>Claim Number:  " & strClaim & "</p><p>We attempted to contact you via phone today and unfortunately were not able to reach you. " & _
									  "Below Is the information we wanted to provide regarding the Fast-Track claim process. " & _
									  "This process requires you to submit documentation directly to our office for review." & _ 
									  "Please feel free to reach out to us with any questions.</p>"
End If

Select Case File.Drawer.Name
	Case "CPCL"
		strCompanyName = "Capitol Preferred Insurance"
	Case "SFCL"
		strCompanyName = "Southern Fidelity Insurance"
	Case "SPCC"
		strCompanyName = "Southern Fidelity Property & Casualty"
	Case Else
		strCompanyName = "Black Pearl Adjusting"
End Select

Logger.Info "New_Loss_Contact_Email - strEmail = " & strEmail
'strEmail = "dbernath@pmains.com"
Logger.Info "New_Loss_Contact_Email - strEmail = " & strEmail

strSubject = strCompanyName & " " & strSubject & " - " & strClaim
strFrom = "Black Pearl Adjusting<estimates@blackpearladjusting.com>"
strTo = strEmail
strBcc = "ir5claims@pmains.com"
strMessage = "<HTML><BODY>" & strImage & _
						strEmailGreeting & "<p>Please provide to us at <a href='mailto:estimates@blackpearladjusting.com?subject=" & strCompanyName & " Claim " & strClaim &" - " & Task.File.Name & "'>estimates@blackpearladjusting</a> within thirty (30) days the following:</p>" & _
						"<p><ul><li>Two (2) estimates for each type of damage. For example: two roof estimaets, two interior estiamtes, two estimates that cover all damages, etc.</li>" & _
						"<li>The estimes should be of like kind and quality materials.</li>" & _
						"<li>Clear photos of the damage. Roof damage may be photographed from the ground if there is visible damage to shingles or tiles.</li>" & _
						"<li>Signed work aughorizations, if applicable.</li></ul></p>" & _
						"<p><u>Once we are in receipt of your documentation, we will review your policy for coverage and any exclusions that may apply to your damages. " & _
						"Once a decision has been made, you will receive applicable correspondence in the mail.</u></p>" & _
						"<p>You may contact our office at the number listed below with any questions or concerns you have.</p>" & _
						"<p>Your duties after a loss include, but are not limited to:</p>" & _
						"<p><ol><li>Protect the property from further damage;</li>" & _
						"<li>Make reasonable and necessary repairs to protect the property;</li>" & _
						"</li>Keep an accurate record of repair expenses;<li>" & _
						"<li>Provide us with records, photos, and documents we request and permit us to make copies; and</li>" & _
						"<li>As often as we reasonably require, show the damaged property.</li></ol></p>" & _
						"<p> If you have any information or documentation related to this claim that you would like us to consider, please submit that information promptly <b>to the attention of your claim number (" & strClaim & ")</b> by any of the following methods:</p>" & _ 
						"<p>Email: <a href='mailto:estimates@blackpearladjusting.com?subject=" & strCompanyName & " Claim " & strClaim &" - " & Task.File.Name & "'>estimates@blackpearladjusting</a><br/>" & _
						"Fax: 850-521-3069<br/>" & _
						"<table style='border:0px solid white'><tr><td>Email:</td><a href='mailto:estimates@blackpearladjusting.com?subject=" & strCompanyName & " Claim " & strClaim &" - " & Task.File.Name & "'>estimates@blackpearladjusting</a><td></tr>" & _
						"<tr><td>Fax:</td><td>850-521-3069</td></tr>" & _ 
						"<tr><td>Mail:</td><td>Black Pearl Adjusting</td></tr>" & _
						"<tr><td></td><td>P.O. Box ?????</td></tr>" & _
						"<tr><td></td><td>Tallahassee, FL 32317</td></tr></table>" & _
						"<p>AOB disclaimer - <i>If you are asked to sign any document at all, please read it very carefully to see if the document contains wording like """"<b>Assignment of Benefits</b>"""". " & _
						"If you sign such a document, you may be assigning some or all of your rights under your policy to a contractor and there may be no way for you to get them back. " & _
						"If you sign something like that, there Is also nothing we can do to help you get your rights back. " & _
						"Always be careful to read any document you sign, but read especially carefully any document containing an """"<b>Assignment of Benefits</b>"""" because it may limit the rights you have under your insurance policy.</p>" & _
						"<p>If at any time you would like to have an adjuster sent to your home, you can contact our office to have your claim reassigned. " & _
						"<b>Please remember that neither Black Pearl Adjusting or Capitol Preferred Insurance are responsible for the repairs to your home. You will have to get estimates and hire a contractor to make the necessary repairs.</b></p>" & _
						"<p>If you do have any questions about this or the adjustment of your claim, please do not hesitate to contact us. We look forward to assisting you with this claim.</p>" & _
						"<p>Sincerely</p>" & _
						"<p>Black Pearl Adjusting</p><p>Phone: 850-303-9764</p><p><a href='www.blackpearladjusting.com'>www.blackpearladjusting<a></p>" & _
						"</BODY></HTML>"

 Email strSubject, strFrom, strTo, strCc, strBcc, strMessage
Logger.Info "New_Loss_Contact_Email - End"
End Function

'******************** Black Pearl Workflow - Development *************************
'******************** Hub ***************************************************
'******************** Index File ***********************************************
Function Index_File(strFileNumber, strDrawer)
Logger.Info "Index_File - Start"
Logger.Info "Index_File - File Number - " & strFileNumber
Logger.Info "Index_File - Drawer - " & strDrawer

Dim objDrawer, objFile, strFileName
strFileName = ""

	Set objDrawer = ObjectLookup.FindDrawer(Null,strDrawer)
	Set objFile =  TypesLookup.GetFileType(strDrawer)

File.Index objDrawer, objFile, strFileNumber, strFileName

Logger.Info "Index_File - End"
End Function

'******************** Black Pearl Workflow - Development *************************
'******************** Hub ***************************************************
'******************** Email **************************************************
Function Email(strSubject, strFrom, strTo, strCc, strBcc, strMessage)
Logger.Info "Email - Start"
Logger.Info "Email - Subject - " & strSubject
Logger.Info "Email - From - " & strFrom
Logger.Info "Email - To - " & strTo
Logger.Info "Email - Cc - " & strCc
Logger.Info "Email - Bcc - " & strBcc
Logger.Info "Email - Message - " & strMessage

Set objconn = CreateObject("ADODB.Connection")

 'Send the message using the network (SMTP over the network)

 Set objMessage = CreateObject("CDO.Message")
 objMessage.Subject = strSubject
 objMessage.From = strFrom
 objMessage.To = strTo
 objMessage.Cc = strCc
 objMessage.Bcc = strBcc
 objMessage.HTMLBody = strMessage

Const cdoSendUsingPort = 2

Const cdoBasic = 1 'basic (clear-text) authentication
Const cdoNTLM = 2 'NTLM

'==This section provides the configuration information for the remote SMTP server.

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

'Name or IP of Remote SMTP Server
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtprelay.pma.local"

'Type of authentication, NONE, Basic (Base64 encoded), NTLM
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1

'Your UserID on the SMTP server
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusername") = "pma\internalrelay"

'Your password on the SMTP server
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "Password1$"

'Server port (typically 25)
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

'Use SSL for the connection (False or True)
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False

'Connection Timeout in seconds (the maximum time CDO will try to establish a connection to the SMTP server)
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

objMessage.Configuration.Fields.Update

'==End remote SMTP server configuration section==

objMessage.Send
Logger.Info "Email - End"
End Function

'******************** Black Pearl Workflow - Development *************************
'******************** Hub ***************************************************
'******************** Call Database ********************************************
Function Call_Database(strSQL, strDB, sResult)
Logger.Info "Call_Database - Start"

Logger.Info "Call_Database - strSQL - "& strSQL
Logger.Info "Call_Database - strDB - " & strDB

Dim rs

'Pull data from CDS
Set objconn = CreateObject ("ADODB.Connection")
objconn.ConnectionTimeout=0

objconn.Open strDB

If objconn.state = 0 Then
     Task.GetAttributeObject("ROUTER").Value = "BADCONNECTION"
     sResult = "Bad Connection"
     Logger.Info "Call_Database - Bad Connection"
     Exit Function
End If

Set rs = CreateObject("ADODB.Recordset")

'This connection gets the data from the database
rs.Open strSQL, objconn
If InStr(strSQL, "dbo.IR_UpdateInsuredContact") Or InStr(strSQL, "dbo.IR_UpdateAlternateContact")Or InStr(strSQL,"dbo.CDS_Status_Update") Or InStr(strSQL, "BPA.updateClaimAdjustingStatus") Then
	Logger.Info "Call_Database - No Result"
	sResult = "No Result"
Else
	sResult = rs.Fields(0).Value
	Logger.Info "Call_Database - Result - " & sResult
	rs.close
End If

Logger.Info "Call_Database - End"
End Function

'******************** Black Pearl Workflow - Development *************************
'******************** Hub ***************************************************
'******************** AutoNote ************ **********************************
Function AutoNote(strDrawer, strFile, strFlow, strStep, strUser, strPriority, strTaskDescription, strNote)
Logger.Info "Autonote - Start"
Dim objconn: Set objconn = CreateObject("ADODB.Connection")
'Calculate the date and time for the file note
Dim strsql
Dim strYear:	strYear = Year(Now): 	Logger.Info "AutoNote - Year - " & strYear
Dim strMonth:	strMonth = Month(Now):	Logger.Info "AutoNote - Month - " & strMonth
If Len(strMonth) = "1" Then 
	strMonth = "0" & strMonth
	Logger.Info "AutoNote - Month Modified - " & strMonth
End If
Dim strDay:		strDay = Day(Now): 		Logger.Info "AutoNote - Day - " & strDay
If Len(strDay) = "1" Then 
	strDay = "0" & strDay
	Logger.Info "AutoNote - Day Modified - " & strDay
End If
Dim strDate:	strDate = strYear & strMonth & strDay:			Logger.Info "AutoNote - Date - " & strDate

Dim strHour:	strHour = Hour(Now):		Logger.Info "AutoNote - Hour - " & strHour
If Len(strHour) = "1" Then 
	strHour = "0" & strHour
	Logger.Info "AutoNote - Hour Modified - " & strHour
End If
Dim strMinute:	strMinute = Minute(Now):	Logger.Info "AutoNote - Minute - " & strMinute
If Len(strMinute) = "1" Then 
	strMinute = "0" & strMinute
	Logger.Info "AutoNote - Minute Modified - " & strMinute
End If
Dim strSeconds:	strSeconds = Second(Now):	Logger.Info "AutoNote - Seconds - " & strSeconds
If Len(strSeconds) = "1" Then 
	strSeconds = "0" & strSeconds
	Logger.Info "AutoNote - Seconds Modified - " & strSeconds
End If
Dim strTime:	strTime = strHour & ":" & strMinute & ":" & strSeconds:					Logger.Info "AutoNote - Time - " & strTime
Logger.Info "AutoNote - Drawer - " & strDrawer
Logger.Info "AutoNote - File - " & strFile
Logger.Info "AutoNote - Flow - " & strFlow
Logger.Info "AutoNote - Step - " & strStep
Logger.Info "AutoNote - User - " & strUser
Logger.Info "AutoNote - Task Description - " & strTaskDescription
Logger.Info "AutoNote - Note - " & strNote

objconn.ConnectionTimeout = 0
objconn.Open "PROVIDER=SQLOLEDB;Data Source=IR5DBdev;Initial Catalog=Utility;uid=sa;PWD=cpic0742"
strsql = "exec createImageRightNote_5X '" + strDrawer + "', '" + strFile + "', '" + strUser  + "', '" + strDate + "', '" + strTime + "', '" + strFlow + "','" + strStep + "',Null,'" + strPriority + "',Null,'" + strDrawer + "','" + strTaskDescription + "',Null,'" + strNote + "'"
Logger.Info "AutoNote - strsql - " & strsql
Set objrs = objconn.execute(strsql)
objconn.close

Logger.Info "Autnote - End"
End Function