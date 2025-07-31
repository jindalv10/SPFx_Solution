
$EPApproverSiteURL = "https://https://spo365dev.sharepoint.com//sites/PRANHub"
$NominationSiteURL = "https://https://spo365dev.sharepoint.com//sites/PDnominations"


$ListName= "EP Admin"
$ContentTypeName ="EP Admin"
$ParentContentTypeName ="0x01"
$ColumnGroup="PD Nomination"

Connect-PnPOnline $NominationSiteURL -UseWebLogin

#Create Nomination Site columns
Add-PnPField -DisplayName "Nominee Name" -InternalName "NomineeName" -Type User -Group "PD Nomination"  -Required
Add-PnPField -DisplayName "FinanceUserID" -InternalName "FinanceUserID" -Type Text -Group "PD Nomination"  -Required
Add-PnPField -DisplayName "NomineeDesignation" -InternalName "NomineeDesignation" -Type Text -Group "PD Nomination"
Add-PnPField -DisplayName "Office" -InternalName "NomineeOffice" -Type Text -Group "PD Nomination"
Add-PnPField -DisplayName "Practice" -InternalName "NomineePractice" -Type Text -Group "PD Nomination"
Add-PnPField -DisplayName "Discipline" -InternalName "NomineeDiscipline" -Type Text -Group "PD Nomination"
Add-PnPField -DisplayName "Is Product Person" -InternalName "IsProductPerson" -Type Boolean -Group "PD Nomination"
Add-PnPField -DisplayName "PD Status" -InternalName "PDStatus" -Type Text -Group "PD Nomination"
Add-PnPField -DisplayName "PD Discipline" -InternalName "PDDiscipline" -Type Text -Group "PD Nomination"
Add-PnPField -DisplayName "PD Subcategory" -InternalName "PDSubcategory" -Type Text -Group "PD Nomination"
Add-PnPField -DisplayName "Proficient Language" -InternalName "ProficientLanguage" -Type Text -Group "PD Nomination"
Add-PnPField -DisplayName "Intake Notes" -InternalName "IntakeNotes" -Type Note -Group "PD Nomination"
Add-PnPField -DisplayName "Status" -InternalName "NominationStatus" -Type Text 
Add-PnPField -DisplayName "RP Certification" -InternalName "RPCertification" -Type Boolean -Group "PD Nomination"
Add-PnPField -DisplayName "Submission Date" -InternalName "SubmissionDate" -Type DateTime -Group "PD Nomination"
Add-PnPField -DisplayName "Draft Date" -InternalName "DraftDate" -Type DateTime -Group "PD Nomination"
Add-PnPField -DisplayName "Re Submission Date" -InternalName "ReSubmissionDate" -Type DateTime -Group "PD Nomination"
Add-PnPField -DisplayName "Is Status Granted After 2016" -InternalName "IsStatusGrantedAfter2016" -Type Boolean -Group "PD Nomination"
$FieldXML= "<Field Type='UserMulti' Name='EPNominator' ID='$([GUID]::NewGuid())' DisplayName='EP Nominator(s)' Required ='TRUE' UserSelectionMode='PeopleOnly' Mult='TRUE' Group='PD Nomination' ></Field>"
Add-PnPFieldFromXml $FieldXML

#Add Fields to content type
$parentContentType = Get-PnPContentType -Identity "Item"
$addedContentType = Add-PnPContentType -Name "PD Nomination Details" -Description "Use for capturing PD Nomination Details" -Group "PD Nomination Content Types" -ParentContentType $parentContentType
Add-PnPFieldToContentType -Field "EPNominator" -ContentType "PD Nomination Details"
Add-PnPFieldToContentType -Field "NomineeName" -ContentType "PD Nomination Details"
Add-PnPFieldToContentType -Field "FinanceUserID" -ContentType "PD Nomination Details"
Add-PnPFieldToContentType -Field "NomineeOffice" -ContentType "PD Nomination Details"
Add-PnPFieldToContentType -Field "NomineePractice" -ContentType "PD Nomination Details"
Add-PnPFieldToContentType -Field "NomineeDiscipline" -ContentType "PD Nomination Details"
Add-PnPFieldToContentType -Field "IsProductPerson" -ContentType "PD Nomination Details"
Add-PnPFieldToContentType -Field "PDStatus" -ContentType "PD Nomination Details"
Add-PnPFieldToContentType -Field "PDDiscipline" -ContentType "PD Nomination Details"
Add-PnPFieldToContentType -Field "PDSubcategory" -ContentType "PD Nomination Details"
Add-PnPFieldToContentType -Field "ProficientLanguage" -ContentType "PD Nomination Details"
Add-PnPFieldToContentType -Field "IntakeNotes" -ContentType "PD Nomination Details"
Add-PnPFieldToContentType -Field "NominationStatus" -ContentType "PD Nomination Details"
Add-PnPFieldToContentType -Field "RPCertification" -ContentType "PD Nomination Details"
Add-PnPFieldToContentType -Field "SubmissionDate" -ContentType "PD Nomination Details"
Add-PnPFieldToContentType -Field "DraftDate" -ContentType "PD Nomination Details"
Add-PnPFieldToContentType -Field "ReSubmissionDate" -ContentType "PD Nomination Details"
Add-PnPFieldToContentType -Field "IsStatusGrantedAfter2016" -ContentType "PD Nomination Details"
Add-PnPFieldToContentType -Field "NomineeDesignation" -ContentType "PD Nomination Details"

#Create New PD Nomination Details list
$masterList = "PD Nominations"
New-PnPList -Title $masterList -Template GenericList
Add-PnPContentTypeToList -List $masterList -ContentType "PD Nomination Details" -DefaultContentType

$LookupListID = (Get-PnPList -Identity $masterList).ID
$lookupFieldXML= "<Field Type='Lookup' Name='Nomintaion' ID='$([GUID]::NewGuid())' DisplayName='Nomination' List='$LookupListID' Group='PD Nomination' ShowField='Title,NominationStatus'></Field>"
Add-PnPFieldFromXml $lookupFieldXML

Add-PnPField -DisplayName "Assignee" -InternalName "Assignee" -Type User -Group "PD Nomination"  -Required 
Add-PnPField -DisplayName "Is Employee Agreement Signed" -InternalName "IsEmployeeAgreementSigned" -Type Boolean -Group "PD Nomination"
Add-PnPField -DisplayName "Is Employee Number Updated" -InternalName "IsEmployeeNumberUpdated" -Type Boolean -Group "PD Nomination"
Add-PnPField -DisplayName "Review Date" -InternalName "ReviewDate" -Type DateTime -Group "PD Nomination"
Add-PnPField -DisplayName "Review Notes" -InternalName "ReviewNotes" -Type Note -Group "PD Nomination"
Add-PnPField -DisplayName "Employee Number Reversed Date" -InternalName "EmployeeNumberReversedDate" -Type DateTime -Group "PD Nomination"
Add-PnPField -DisplayName "Is Employee Number Reversed" -InternalName "IsEmployeeNumberReversed" -Type Boolean -Group "PD Nomination"
Add-PnPField -DisplayName "Withdraw Completion Date" -InternalName "WithdrawCompletionDate" -Type DateTime -Group "PD Nomination"

#Create New NominationDetailsByLocalAdmin list
$addedContentType = Add-PnPContentType -Name "Nomination Details By LA" -Description "Use for capturing Nomination Details by LA" -Group "PD Nomination Content Types" -ParentContentType $parentContentType
Add-PnPFieldToContentType -Field "Assignee" -ContentType "Nomination Details By LA"
Add-PnPFieldToContentType -Field "IsEmployeeAgreementSigned" -ContentType "Nomination Details By LA"
Add-PnPFieldToContentType -Field "IsEmployeeNumberUpdated" -ContentType "Nomination Details By LA"
Add-PnPFieldToContentType -Field "ReviewNotes" -ContentType "Nomination Details By LA"
Add-PnPFieldToContentType -Field "ReviewDate" -ContentType "Nomination Details By LA"
Add-PnPFieldToContentType -Field "EmployeeNumberReversedDate" -ContentType "Nomination Details By LA"
Add-PnPFieldToContentType -Field "IsEmployeeNumberReversed" -ContentType "Nomination Details By LA"
Add-PnPFieldToContentType -Field "WithdrawCompletionDate" -ContentType "Nomination Details By LA"
Add-PnPFieldToContentType -Field "Nomintaion" -ContentType "Nomination Details By LA"

$LocalAdminList = "Nomination Details By Local Admin"
New-PnPList -Title $LocalAdminList -Template GenericList
Add-PnPContentTypeToList -List $LocalAdminList -ContentType "Nomination Details By LA" -DefaultContentType

#Create Site columns and Contnet type for GCS Leagal
Add-PnPField -DisplayName "Is Employee Agreement Signed By CEO" -InternalName "IsEmpAgreementSignedByCEO" -Type Boolean -Group "PD Nomination"
Add-PnPField -DisplayName "Is Saved On Local Drive" -InternalName "IsSavedOnLocalDrive" -Type Boolean -Group "PD Nomination"
Add-PnPField -DisplayName "Reviewer" -InternalName "Reviewer" -Type User -Group "PD Nomination"  -Required 

#Create New NominationDetailsByLegal list
$addedContentType = Add-PnPContentType -Name "Nomination Details By Legal" -Description "Use for capturing Nomination Details by Legal" -Group "PD Nomination Content Types" -ParentContentType $parentContentType
Add-PnPFieldToContentType -Field "Reviewer" -ContentType "Nomination Details By Legal"
Add-PnPFieldToContentType -Field "IsEmpAgreementSignedByCEO" -ContentType "Nomination Details By Legal"
Add-PnPFieldToContentType -Field "IsSavedOnLocalDrive" -ContentType "Nomination Details By Legal"
Add-PnPFieldToContentType -Field "ReviewDate" -ContentType "Nomination Details By Legal"
Add-PnPFieldToContentType -Field "Nomintaion" -ContentType "Nomination Details By Legal"

$GcsLegalList = "Nomination Details By GCS Legal"
New-PnPList -Title $GcsLegalList -Template GenericList
Add-PnPContentTypeToList -List $GcsLegalList -ContentType "Nomination Details By Legal" -DefaultContentType

#Create Site columns and Content type for QC
Add-PnPField -DisplayName "Sent To Sc Date" -InternalName "SentToScDate" -Type DateTime -Group "PD Nomination"
Add-PnPField -DisplayName "Sent For More Details" -InternalName "SentForMoreDetails" -Type DateTime -Group "PD Nomination"
Add-PnPField -DisplayName "Withdrawn Date" -InternalName "WithdrawnDate" -Type DateTime -Group "PD Nomination"
Add-PnPField -DisplayName "Sent To PTPAC Date" -InternalName "SentToPTPACDate" -Type DateTime -Group "PD Nomination"
Add-PnPField -DisplayName "Notifcation Recipient" -InternalName "NotifcationRecipient" -Type Choice -Group "PD Nomination" -Choices "Nominee,Nominator,Both"
Add-PnPField -DisplayName "Granted" -InternalName "Granted" -Type DateTime -Group "PD Nomination"
Add-PnPField -DisplayName "End Date" -InternalName "NominationEndDate" -Type DateTime -Group "PD Nomination"
Add-PnPField -DisplayName "Additional Reviewer" -InternalName "AdditionalReviewer" -Type User -Group "PD Nomination"
Add-PnPField -DisplayName "QC Status" -InternalName "QCStatus" -Type Text -Group "PD Nomination"
Add-PnPField -DisplayName "Reviewer Assignment Date" -InternalName "ReviewerAssignmentDate" -Type DateTime -Group "PD Nomination"
Add-PnPField -DisplayName "Anyone Else" -InternalName "AnyoneElse" -Type Text -Group "PD Nomination"
Add-PnPField -DisplayName "AddPracticeDirector" -InternalName "AddPracticeDirector" -Type Boolean -Group "PD Nomination"
Add-PnPField -DisplayName "NominationEndDate" -InternalName "NominationEndDate" -Type DateTime -Group "PD Nomination"


#Create New NominationDetailsByQC list

$addedContentType = Add-PnPContentType -Name "Nomination Details By QC" -Description "Use for capturing Nomination Details by QC" -Group "PD Nomination Content Types" -ParentContentType $parentContentType
Add-PnPFieldToContentType -Field "ReviewNotes" -ContentType "Nomination Details By QC"
Add-PnPFieldToContentType -Field "Reviewer" -ContentType "Nomination Details By QC"
Add-PnPFieldToContentType -Field "AdditionalReviewer" -ContentType "Nomination Details By QC"
Add-PnPFieldToContentType -Field "ReviewDate" -ContentType "Nomination Details By QC"
Add-PnPFieldToContentType -Field "SentToScDate" -ContentType "Nomination Details By QC"
Add-PnPFieldToContentType -Field "DraftDate" -ContentType "Nomination Details By QC"
Add-PnPFieldToContentType -Field "SentForMoreDetails" -ContentType "Nomination Details By QC"
Add-PnPFieldToContentType -Field "WithdrawnDate" -ContentType "Nomination Details By QC"
Add-PnPFieldToContentType -Field "SentToPTPACDate" -ContentType "Nomination Details By QC"
Add-PnPFieldToContentType -Field "NotifcationRecipient" -ContentType "Nomination Details By QC"
Add-PnPFieldToContentType -Field "Nomintaion" -ContentType "Nomination Details By QC"
Add-PnPFieldToContentType -Field "Granted" -ContentType "PD Nomination Details By QC"
Add-PnPFieldToContentType -Field "NominationEndDate" -ContentType "PD Nomination Details By QC"
Add-PnPFieldToContentType -Field "QCStatus" -ContentType "PD Nomination Details By QC"
Add-PnPFieldToContentType -Field "ReviewerAssignmentDate" -ContentType "Nomination Details By QC"
Add-PnPFieldToContentType -Field "AnyoneElse" -ContentType "Nomination Details By QC"
Add-PnPFieldToContentType -Field "AddPracticeDirector" -ContentType "Nomination Details By QC"
Add-PnPFieldToContentType -Field "NominationEndDate" -ContentType "Nomination Details By QC"

$QCList = "Nomination Details By QC"
New-PnPList -Title $QCList -Template GenericList
Add-PnPContentTypeToList -List $QCList -ContentType "Nomination Details By QC" -DefaultContentType


#Create Site columns and Content type for PTPAC
Add-PnPField -DisplayName "PTPAC Chair" -InternalName "PTPACChair" -Type User -Group "PD Nomination"  -Required 
Add-PnPField -DisplayName "Recommendation" -InternalName "Recommendation" -Type Note -Group "PD Nomination"
Add-PnPField -DisplayName "Review Due Date" -InternalName "ReviewDueDate" -Type DateTime -Group "PD Nomination"
Add-PnPField -DisplayName "Recommendation Sent Date" -InternalName "RecommendationSentDate" -Type DateTime -Group "PD Nomination"
Add-PnPField -DisplayName "Ptpac Status" -InternalName "PtpacStatus" -Type Text -Group "PD Nomination"

#$LookupListID = (Get-PnPList -Identity $QCList).ID
#$lookupFieldXML= "<Field Type='Lookup' Name='QCDetails' ID='$([GUID]::NewGuid())' DisplayName='QCDetails' List='$LookupListID' Group='PD Nomination' ShowField='Title,QCStatus'></Field>"
#Add-PnPFieldFromXml $lookupFieldXML

#Create New NominationDetailsByPTPAC list
$addedContentType = Add-PnPContentType -Name "Nomination Details By PTPAC" -Description "Use for capturing Nomination Details by PTPAC" -Group "PD Nomination Content Types" -ParentContentType $parentContentType
Add-PnPFieldToContentType -Field "Reviewer" -ContentType "Nomination Details By PTPAC"
Add-PnPFieldToContentType -Field "PTPACChair" -ContentType "Nomination Details By PTPAC"
Add-PnPFieldToContentType -Field "ReviewDueDate" -ContentType "Nomination Details By PTPAC"
Add-PnPFieldToContentType -Field "ReviewDate" -ContentType "Nomination Details By PTPAC"
Add-PnPFieldToContentType -Field "RecommendationSentDate" -ContentType "Nomination Details By PTPAC"
Add-PnPFieldToContentType -Field "ReviewerAssignmentDate" -ContentType "Nomination Details By PTPAC"
Add-PnPFieldToContentType -Field "Nomintaion" -ContentType "Nomination Details By PTPAC"
Add-PnPFieldToContentType -Field "PtpacStatus" -ContentType "PD Nomination Details By PTPAC"
#Add-PnPFieldToContentType -Field "QCDetails" -ContentType "Nomination Details By PTPAC"

$PtpacList = "Nomination Details By PTPAC"
New-PnPList -Title $PtpacList -Template GenericLists
Add-PnPContentTypeToList -List $PtpacList -ContentType "Nomination Details By PTPAC" -DefaultContentType

#Create Site columns and Content type for Attachments
Add-PnPField -DisplayName "Attachment Type" -InternalName "AttachmentType" -Type Choice -Group "PD Nomination" -Choices "Nomination form,Reference,PRPR"

$parentContentType = Get-PnPContentType -Identity "Document Set"
$addedContentType = Add-PnPContentType -Name "Nomination Attachments" -Description "Use for capturing Nomination attachments" -Group "PD Nomination Content Types" -ParentContentType $parentContentType
Add-PnPFieldToContentType -Field "AttachmentType" -ContentType "Nomination Attachments"
Add-PnPFieldToContentType -Field "Nomintaion" -ContentType "Nomination Attachments"
$AttachmentList = "Nomination Attachments"
New-PnPList -Title $AttachmentList -Template DocumentLibrary
Add-PnPContentTypeToList -List $AttachmentList -ContentType "Nomination Attachments" -DefaultContentType

#Create Site columns and Content type for Email template

Add-PnPField -DisplayName "Email To" -InternalName "emailTo" -Type Text -Group "PD Nomination"
Add-PnPField -DisplayName "Email CC" -InternalName "emailCC" -Type Text -Group "PD Nomination"
Add-PnPField -DisplayName "Email Subject" -InternalName "emailSub" -Type Text -Group "PD Nomination"
Add-PnPField -DisplayName "Email Body" -InternalName "emailBody" -Type Note -Group "PD Nomination"
Add-PnPField -DisplayName "Email Title" -InternalName "emailTitle" -Type Text -Required -Group "PD Nomination"
Add-PnPField -DisplayName "Notification PD Discipline" -InternalName "notificationPDDiscipline" -Type Choice -Choices "Joint","Health","Life and Financial Services","Employee Benefits","Property and Casualty","Global Corporate Services" -Group "PD Nomination"
Add-PnPField -DisplayName "Notification PD Status" -InternalName "notificationPDStatus" -Type Choice -Choices "Approved Professional","Limited Signature Authority","Signature Authority","Qualified Reviewer" -Group "PD Nomination"


$parentContentType = Get-PnPContentType -Identity "Item"
$addedContentType = Add-PnPContentType -Name "Nomination Email template" -Description "Use for storing Nomination email configuration" -Group "PD Nomination Content Types" -ParentContentType $parentContentType
Add-PnPFieldToContentType -Field "emailTo" -ContentType "Nomination Email template"
Add-PnPFieldToContentType -Field "emailCC" -ContentType "Nomination Email template"
Add-PnPFieldToContentType -Field "emailBody" -ContentType "Nomination Email template"
Add-PnPFieldToContentType -Field "emailTitle" -ContentType "Nomination Email template"
Add-PnPFieldToContentType -Field "notificationPDDiscipline" -ContentType "Nomination Email template"
Add-PnPFieldToContentType -Field "notificationPDStatus" -ContentType "Nomination Email template"

$EmailList = "Nomination Notifications"
New-PnPList -Title $EmailList -Template GenericList
Add-PnPContentTypeToList -List $EmailList -ContentType "Nomination Email template" -DefaultContentType

<#
Add-PnPField -DisplayName "Requestor" -InternalName "Requestor" -Type User -Group "PD Nomination"  -Required 
Add-PnPField -DisplayName "EP" -InternalName "EPApprover" -Type User -Group "PD Nomination"  -Required 
Add-PnPField -DisplayName "Notes" -InternalName "RequestorNotes" -Type Note -Group "PD Nomination"   
Add-PnPField -DisplayName "Request Status" -InternalName "RequestStatus" -Type Text -Group "PD Nomination"   
Add-PnPField -DisplayName "Response Date" -InternalName "ResponseDate" -Type DateTime -Group "PD Nomination"   

$addedContentType = Add-PnPContentType -Name "EP Admin" -Description "Use for approval by EP for Admin request" -Group "PD Nomination Content Types" -ParentContentType $parentContentType
Add-PnPFieldToContentType -Field "Requestor" -ContentType "EP Admin"
Add-PnPFieldToContentType -Field "EPApprover" -ContentType "EP Admin"
Add-PnPFieldToContentType -Field "RequestorNotes" -ContentType "EP Admin"
Add-PnPFieldToContentType -Field "RequestStatus" -ContentType "EP Admin"
Add-PnPFieldToContentType -Field "ResponseDate" -ContentType "EP Admin"

$EPAdminList = "EP Admin"
New-PnPList -Title $EPAdminList -Template GenericList
Add-PnPContentTypeToList -List $EPAdminList -ContentType "EP Admin" -DefaultContentType
#>

$FieldXML= "<Field Type='UserMulti' Name='AuthorizedPTPAC' ID='$([GUID]::NewGuid())' DisplayName='Authorized PTPAC' Required ='TRUE' UserSelectionMode='PeopleOnly' Mult='TRUE' Group='PD Nomination' ></Field>"
Add-PnPFieldFromXml $FieldXML
$FieldXML= "<Field Type='UserMulti' Name='AuthorizedQC' ID='$([GUID]::NewGuid())' DisplayName='Authorized QC' Required ='TRUE' UserSelectionMode='PeopleOnly' Mult='TRUE' Group='PD Nomination' ></Field>"
Add-PnPFieldFromXml $FieldXML

$parentContentType = Get-PnPContentType -Identity "Item"
$addedContentType = Add-PnPContentType -Name "Nomination Reviewers" -Description "Use for storing QC and PTPAC reviewers" -Group "PD Nomination Content Types" -ParentContentType $parentContentType
Add-PnPFieldToContentType -Field "AuthorizedPTPAC" -ContentType "Nomination Reviewers"
Add-PnPFieldToContentType -Field "AuthorizedQC" -ContentType "Nomination Reviewers"
Add-PnPFieldToContentType -Field "PDDiscipline" -ContentType "Nomination Reviewers"

$NominationReviewers = "Nomination Reviewers"
New-PnPList -Title $NominationReviewers -Template GenericList
Add-PnPContentTypeToList -List $NominationReviewers -ContentType "Nomination Reviewers" -DefaultContentType

Try {
    #Connect to PnP Online
    Connect-PnPOnline -Url $EPApproverSiteURL -UseWebLogin
	
	Write-host "########## Provisioning EP Approver Script Start ########################" -f Yellow
	
	$list = Get-PnPList -Identity $ListName
	If($list.Title -ne $ListName) {  
		
		Write-Host "########## List Creation Started ########################"  -f cyan
		
		New-PnPList -Title $ListName -Url "lists/$ListName" -Template GenericList
		
		
		# Create SiteColumns 
		
	    $Field1 = Add-PnPField  -DisplayName "Requestor:" -InternalName "Requestor" -Type User -Required -AddToDefaultView -Group $ColumnGroup
		$Field2 = Add-PnPField  -DisplayName "EP Approver:" -InternalName "EPApprover" -Type User -Required -AddToDefaultView -Group $ColumnGroup
		$Field3 = Add-PnPField  -DisplayName "EP Notes" -InternalName "EPNotes" -Type Note -AddToDefaultView -Group $ColumnGroup
		$Field4 = Add-PnPField  -DisplayName "Decision Date" -InternalName "DecisionDate" -Type DateTime -AddToDefaultView -Group $ColumnGroup
		$Field5 = Add-PnPField  -DisplayName "Status" -InternalName "Status" -Type Choice -AddToDefaultView -Choices "New","Pending","Approved","Rejected" -Group $ColumnGroup
		$Field6 = Add-PnPField  -DisplayName "ApprovalLogs" -InternalName "ApprovalLogs" -Type Note -AddToDefaultView -Group $ColumnGroup
		
		
		# Create content type
		$ParentContentType = Get-PnPContentType -Identity $ParentContentTypeName
		$EPApproverCT = Add-PnPContentType -Name $ContentTypeName -Description "Use for approval by EP for Admin request" -Group "PD Nomination Content Types" -ParentContentType $ParentContentType
		
		
		
		#Link site Column to content type
        Add-PnPFieldToContentType -Field $Field1 -ContentType $EPApproverCT 
        Add-PnPFieldToContentType -Field $Field2 -ContentType $EPApproverCT
		Add-PnPFieldToContentType -Field $Field3 -ContentType $EPApproverCT
        Add-PnPFieldToContentType -Field $Field4 -ContentType $EPApproverCT
		Add-PnPFieldToContentType -Field $Field5 -ContentType $EPApproverCT
        Add-PnPFieldToContentType -Field $Field6 -ContentType $EPApproverCT
		
		
		
	
		# Add content type to list
        $ListName = Get-PnPList "/lists/$ListName"
        Add-PnPContentTypeToList -List $ListName -ContentType $EPApproverCT -DefaultContentType
		
		
		
		Write-host "New List ,Site Columns and ContentType Added Successfully.!" -f Green
	}  
	else {  
		Write-host "########## Error: List Already Exist.! ########################" -f Red
	}  
	Write-host "########## Provisioning Script End ########################" -f Yellow
    }
catch 
{
    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
}

#Add Nomination Reviewers content Type