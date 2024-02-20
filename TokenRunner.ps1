
##################################################
# Script  : Runner-qr.ps1
# Author  : Noel Williams
# Updated : 22/03/23   
# Input   : Null
# Output  : converted QR Image or txt file of sdtid file is place in converted dir
################################################## 

$intro = @"
    

░██████╗░█████╗░███╗░░██╗░█████╗░████████╗██╗░░░██╗███╗░░░███╗  ████████╗░█████╗░██╗░░██╗███████╗███╗░░██╗
██╔════╝██╔══██╗████╗░██║██╔══██╗╚══██╔══╝██║░░░██║████╗░████║  ╚══██╔══╝██╔══██╗██║░██╔╝██╔════╝████╗░██║
╚█████╗░███████║██╔██╗██║██║░░╚═╝░░░██║░░░██║░░░██║██╔████╔██║  ░░░██║░░░██║░░██║█████═╝░█████╗░░██╔██╗██║
░╚═══██╗██╔══██║██║╚████║██║░░██╗░░░██║░░░██║░░░██║██║╚██╔╝██║  ░░░██║░░░██║░░██║██╔═██╗░██╔══╝░░██║╚████║
██████╔╝██║░░██║██║░╚███║╚█████╔╝░░░██║░░░╚██████╔╝██║░╚═╝░██║  ░░░██║░░░╚█████╔╝██║░╚██╗███████╗██║░╚███║
╚═════╝░╚═╝░░╚═╝╚═╝░░╚══╝░╚════╝░░░░╚═╝░░░░╚═════╝░╚═╝░░░░░╚═╝  ░░░╚═╝░░░░╚════╝░╚═╝░░╚═╝╚══════╝╚═╝░░╚══╝

░█████╗░░█████╗░███╗░░██╗██╗░░░██╗███████╗██████╗░████████╗███████╗██████╗░
██╔══██╗██╔══██╗████╗░██║██║░░░██║██╔════╝██╔══██╗╚══██╔══╝██╔════╝██╔══██╗
██║░░╚═╝██║░░██║██╔██╗██║╚██╗░██╔╝█████╗░░██████╔╝░░░██║░░░█████╗░░██████╔╝
██║░░██╗██║░░██║██║╚████║░╚████╔╝░██╔══╝░░██╔══██╗░░░██║░░░██╔══╝░░██╔══██╗
╚█████╔╝╚█████╔╝██║░╚███║░░╚██╔╝░░███████╗██║░░██║░░░██║░░░███████╗██║░░██║
░╚════╝░░╚════╝░╚═╝░░╚══╝░░░╚═╝░░░╚══════╝╚═╝░░╚═╝░░░╚═╝░░░╚══════╝╚═╝░░╚═╝
"@

$outro = @"
    
███████╗██╗░░██╗██╗████████╗██╗███╗░░██╗░██████╗░
██╔════╝╚██╗██╔╝██║╚══██╔══╝██║████╗░██║██╔════╝░
█████╗░░░╚███╔╝░██║░░░██║░░░██║██╔██╗██║██║░░██╗░
██╔══╝░░░██╔██╗░██║░░░██║░░░██║██║╚████║██║░░╚██╗
███████╗██╔╝╚██╗██║░░░██║░░░██║██║░╚███║╚██████╔╝
╚══════╝╚═╝░░╚═╝╚═╝░░░╚═╝░░░╚═╝╚═╝░░╚══╝░╚═════╝░  
"@

#Email script that takes converted token populates an introduction email with instruction on setup etc, just need to input users email
function emailerFunction {

    $val = Get-ChildItem -Path C:\temp\Converter\converted\* 

    if($val -eq $null) {
        Write-Host "`n`nConverted folder empty, please add txt or image files to the converted folder folder and run the script again.`nExiting`n`n"
        exit
    } else {
        Write-Host "`n`nFiles Found!`n"
    }

    foreach ($i in $val) {

        
        # Create a new Outlook application object
        $outlook = New-Object -ComObject Outlook.Application
        
    	$fName = $i.ToString().Split("\")[4]
        $name = $fName.Split("_")[0]
		Write-Host "`t$i"

        # Set the required variables
        $subject = "Welcome To The Club $name"
        $body = "<HTML><p>Dear $name,</p><p>`n`nI am writing to inform you that we have attached your L4 Club Token and activation instructions to this email. This token is an essential security measure that helps protect your account from unauthorized access.</p><p>To activate your Club Token, please follow the instructions attached to this email carefully. The instructions provide step-by-step guidance on how to download and install the SecureId software, as well as how to activate your Token.</p><p>If you encounter any issues during the activation process, please do not hesitate to contact our Service Desk. <br/>Our team is available to provide you with support and assistance to ensure smooth Token activation.`n`nThank you for your cooperation in this matter.`n`n</p><HTML>"
        #add a way to create multiple depending on files in folder
        $imagePath = "$i"
        #change path to doc of your liking
        $docPath = "C:\temp\Converter\docs\Club_help_doc.docx"
        #change path to htm file for email signiture
        $signaturePath = "C:\temp\Converter\docs\Club.htm"

        [string]$signature = Get-Content $signaturePath

        # Create a new email message
        $email = $outlook.CreateItem(0)


        # Add the recipient, subject, and body
        $email.To = $null 
        $email.Subject = $subject
        #$email.Body = $body 

        # Add the image file as an attachment
        $attachment1 = $email.Attachments.Add($imagePath)

        # Add the Word document as an attachment
        $attachment2 = $email.Attachments.Add($docPath)

        # Add the signature
        [string]$signature = Get-Content $signaturePath
        $email.HTMLBody += $body + $signature 

        Write-Host "`t$name email created and opening.`n"

        # Display the email for review
        $email.Display()

	}

    Write-Host "`n``nEmailer End`n`n"

}


function outroMes {
    Write-Host "Author - Noel Williams$outro"
}

#change this to match the path to script (Right-click script > properties > copy and paste location)
$mainPath = "C:\temp\Converter"

function getFileName {
    
	$val = Get-ChildItem -Path $mainPath\token\* -Include "*.sdtid"
	$fileNameArray = @()
    
	if($val -eq $null) {
        Write-Host "`n`nToken folder empty, please add sdtid files to the token folder and run the script again.`n`n"
        outroMes
        exit
    }

	Write-Host "`nPlease remove files that you do not want converted from the token folder.`n`nThe files listed below will be converted:`n"
	foreach ($i in $val) {
        
    	$name = $i.ToString().Split("\")[4].Split(".")[0]
		$fileNameArray += $name
		Write-Host "`tFile: $name"
	}
 	
	$ans = Read-Host "`nWould you like to continue with the conversion process? (y\n)"

	if( $ans -eq 'y') {
		return $fileNameArray
	} else {
		return 'false'
	}
	
	
}

function moveCompletedFilesToCompleted() {
    
    $CFiles = Get-ChildItem -Path $mainPath\token\* -Include "*.sdtid"
    foreach ($file in $CFiles) {
        if (Test-Path -Path $mainPath\completed\$name.sdtid){
            Remove-Item $mainPath\completed\$name.sdtid
        }
        
	    Move-Item -Path $file -Destination $mainPath\completed\$name.sdtid
    }

}

function deviceInput {
    param (
    [Parameter()] [string] $file
    )

    #Get file contents and perform stearch for the device type
    $defDevType = ''
    $data = Get-Content .\token\$file.sdtid

    For ($i=0; $i -le $data.Length; $i++) {
        if($data[$i] -like "*DeviceTypeFamily*" ) {
            $defDevType = $data[$i].Trim().Replace("<DeviceTypeFamily>", "").Replace("</DeviceTypeFamily>", "").ToLower()
            Write-Host "`n`nDevice Type: $($defDevType.ToUpper())"
        } 
    }

    if($defDevType -eq 'android') {
        return 1
    } elseif($defDevType -eq 'ios') {
        return 2
    } else {
    #incase there is no match
        return 0
    }
    
}

#function to check if user wants to populate email with data and open for sending to new user
    function createEmail() {
        $emailRunner = Read-Host "Run email script on converted token?(y,n)"
        if($emailRunner -ne 'y' -AND $emailRunner -ne 'n') {
            Write-Host "`n`nEnter a valid answer to continue (y,n)`n`n"
            emailScript
        } else {
            if($emailRunner -eq 'y') {
                Write-Host "`nRunning Emailer Script`n`n"
                emailerFunction
            } elseif($emailRunner -eq 'n') {
                Write-Host "`nFine, I can see help is not needed here.`n`n"
                
            }
        }

    }

# --------------------------------------------- Program start ---------------------------------------------

Write-Host "`n`n$intro`n"

$res = getFileName
$fileName = $res

while ( $res -eq 'false' ) {
	if($res -ne 'false') {
		Break
	}else{

		Write-Host "`nMake Changes and rerun.`n"
        outroMes
		exit
	}
}

try {
    foreach ($file in $fileName) {
                $devType = deviceInput -file $file
                #always picks to convert to QR not txt file for ease of use.
                $outputType = 2

                #main paths for files
                $jarPath = "$mainPath\dependancies\TokenConverter.jar"
                $inputPath = "$mainPath\token\$file.sdtid"
                $outputPath = $(if($outputType -eq 2) { "$mainPath\converted\$file.jpeg"} else { "$mainPath\converted\$file.txt"})
                #Checks devType and inputs flag accordingly and check
                $execTokenConverterJar = "java -jar $jarPath $inputPath $(if($devType -eq 1) {"-android"} else {"-ios"}) -o $outputPath $(if($outputType -eq 2){"-qr"}) -d 3"
                #incase no match skip.

                if($defDevType -eq 0) {
                    continue;
                }

                function convertCompleteMessage() {
                    Write-Host $(If ($picked -eq 2) {"`t$file QR image created."} Else {"`t$file text file created."})
                }

                Write-Host "`nProcessing....`n"

                #Dynamically modifies execution string depending on devType, outputType and outputPath
		        Invoke-Expression $execTokenConverterJar
		        convertCompleteMessage
	}

	
	Write-Host "`n`nAll Tokens Converted.`nMoving .sdtid files to completed folder.`n"
    moveCompletedFilesToCompleted
    
    Write-Host "`Moving Complete`n`n"
    createEmail
    outroMes
    

} catch {
 	$Error[0]
}
