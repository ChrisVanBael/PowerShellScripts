# POWERPOINT
# https://gist.github.com/ap0llo/05cef76e3c4040ee924c4cfeef3f0b40
# https://www.jaapbrasser.com/updated-events-github-repository-convert-pptx-to-pdf/
# https://blogs.technet.microsoft.com/heyscriptingguy/2007/07/27/how-can-i-configure-powerpoint-to-print-handouts-instead-of-slides/
# https://social.technet.microsoft.com/Forums/windowsserver/en-US/3d67094a-51a1-4e14-a07f-e1d2edc5a835/automating-printing-of-powerpoint-presentations?forum=winserverpowershell

# Word
# https://social.technet.microsoft.com/Forums/ie/en-US/445b2429-e33c-4ce0-9d64-dd31422571bf/powershell-script-convert-doc-to-pdf?forum=winserverpowershell



# Powershell script to export Powerpoint Presentations to pdf using the Powerpoint COM API
# Based on a VB script with the same purpose
# http://superuser.com/questions/641471/how-can-i-automatically-convert-powerpoint-to-pdf

function Export-Presentation($inputFile, $outputFile)
{
	# Load Powerpoint Interop Assembly
	[Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Powerpoint") > $null
	[Reflection.Assembly]::LoadWithPartialname("Office") > $null

	$msoFalse =  [Microsoft.Office.Core.MsoTristate]::msoFalse
	$msoTrue =  [Microsoft.Office.Core.MsoTristate]::msoTrue

	$ppFixedFormatIntentScreen = [Microsoft.Office.Interop.PowerPoint.PpFixedFormatIntent]::ppFixedFormatIntentScreen # Intent is to view exported file on screen.
	$ppFixedFormatIntentPrint =  [Microsoft.Office.Interop.PowerPoint.PpFixedFormatIntent]::ppFixedFormatIntentPrint  # Intent is to print exported file.

	$ppFixedFormatTypeXPS = [Microsoft.Office.Interop.PowerPoint.PpFixedFormatType]::ppFixedFormatTypeXPS  # XPS format
	$ppFixedFormatTypePDF = [Microsoft.Office.Interop.PowerPoint.PpFixedFormatType]::ppFixedFormatTypePDF  # PDF format

	$ppPrintHandoutVerticalFirst = 1   # Slides are ordered vertically, with the first slide in the upper-left corner and the second slide below it.
	$ppPrintHandoutHorizontalFirst = 2 # Slides are ordered horizontally, with the first slide in the upper-left corner and the second slide to the right of it.

	$ppPrintOutputSlides = 1               # Slides
	$ppPrintOutputTwoSlideHandouts = 2     # Two Slide Handouts
	$ppPrintOutputThreeSlideHandouts = 3   # Three Slide Handouts
	$ppPrintOutputSixSlideHandouts = 4     # Six Slide Handouts
	$ppPrintOutputNotesPages = 5           # Notes Pages
	$ppPrintOutputOutline = 6              # Outline
	$ppPrintOutputBuildSlides = 7          # Build Slides
	$ppPrintOutputFourSlideHandouts = 8    # Four Slide Handouts
	$ppPrintOutputNineSlideHandouts = 9    # Nine Slide Handouts
	$ppPrintOutputOneSlideHandouts = 10    # Single Slide Handouts

	$ppPrintAll = 1            # Print all slides in the presentation.
	$ppPrintSelection = 2      # Print a selection of slides.
	$ppPrintCurrent = 3        # Print the current slide from the presentation.
	$ppPrintSlideRange = 4     # Print a range of slides.
	$ppPrintNamedSlideShow = 5 # Print a named slideshow.

	$ppShowAll = 1             # Show all.
	$ppShowNamedSlideShow = 3  # Show named slideshow.
	$ppShowSlideRange = 2      # Show slide range.

	
	# start Powerpoint
	$application = New-Object "Microsoft.Office.Interop.Powerpoint.ApplicationClass" 

	# Make sure inputFile is an absolute path
	$inputFile = Resolve-Path $inputFile
	#$outputFile = Resolve-Path $outputFile
	$outputFile = [System.IO.Path]::ChangeExtension($inputFile, ".pdf")
	
	$application.Visible = $msoTrue
	$presentation = $application.Presentations.Open($inputFile, $msoTrue, $msoFalse, $msoFalse)
	$printOptions = $presentation.PrintOptions
	$range = $printOptions.Ranges.Add(1,$presentation.Slides.Count) 
	$printOptions.RangeType = $ppShowAll
	
	# export presentation to pdf
	$presentation.ExportAsFixedFormat($outputFile, $ppFixedFormatTypePDF, $ppFixedFormatIntentPrint, $msoTrue, $ppPrintHandoutHorizontalFirst, $ppPrintOutputTwoSlideHandouts, $msoFalse, $range, $ppPrintAll, "Slideshow Name", $False, $False, $False, $False, $False)
	
	$presentation.Close()
	$presentation = $null
	
	if($application.Windows.Count -eq 0)
	{
		$application.Quit()
	}
	
	$application = $null
	
	# Make sure references to COM objects are released, otherwise powerpoint might not close
	# (calling the methods twice is intentional, see https://msdn.microsoft.com/en-us/library/aa679807(office.11).aspx#officeinteroperabilitych2_part2_gc)
	[System.GC]::Collect();
	[System.GC]::WaitForPendingFinalizers();
	[System.GC]::Collect();
	[System.GC]::WaitForPendingFinalizers();

}


function Export-WordDocument($inputFile, $outputFile)
{
	# Load Word Interop Assembly
	[Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Word") > $null
	[Reflection.Assembly]::LoadWithPartialname("Office") > $null
    
    $def = [Type]::Missing
    $msoFalse =  [Microsoft.Office.Core.MsoTristate]::msoFalse
	$msoTrue =  [Microsoft.Office.Core.MsoTristate]::msoTrue

	$wdExportFormatXPS = [Microsoft.Office.Interop.Word.WdExportFormat]::wdExportFormatXPS  # XPS format
	$wdExportFormatPDF = [Microsoft.Office.Interop.Word.WdExportFormat]::wdExportFormatPDF  # PDF format

	$wdExportOptimizeForScreen = [Microsoft.Office.Interop.Word.WdExportOptimizeFor]::wdExportOptimizeForScreen # Intent is to view exported file on screen.
	$wdExportOptimizeForPrint =  [Microsoft.Office.Interop.Word.WdExportOptimizeFor]::wdExportOptimizeForPrint  # Intent is to print exported file.

	$wdExportAllDocument = [Microsoft.Office.Interop.Word.WdExportRange]::wdExportAllDocument  # Exports the entire document
    $wdExportCurrentPage = [Microsoft.Office.Interop.Word.WdExportRange]::wdExportCurrentPage  # Exports the current page
    $wdExportFromTo =      [Microsoft.Office.Interop.Word.WdExportRange]::wdExportFromTo       # Exports the contents of a range using the starting and ending positions
    $wdExportSelection =   [Microsoft.Office.Interop.Word.WdExportRange]::wdExportSelection    # Exports the contents of the current selection

    $wdExportDocumentContent =    [Microsoft.Office.Interop.Word.WdExportItem]::wdExportDocumentContent    # Exports the document without markup
    $wdExportDocumentWithMarkup = [Microsoft.Office.Interop.Word.WdExportItem]::wdExportDocumentWithMarkup # Exports the document with markup

    $wdExportCreateHeadingBookmarks = [Microsoft.Office.Interop.Word.WdExportCreateBookmarks]::wdExportCreateHeadingBookmarks # Create a bookmark in the exported document for each Microsoft Word heading
    $wdExportCreateNoBookmarks =      [Microsoft.Office.Interop.Word.WdExportCreateBookmarks]::wdExportCreateNoBookmarks      # Do not create bookmarks in the exported document
    $wdExportCreateWordBookmarks =    [Microsoft.Office.Interop.Word.WdExportCreateBookmarks]::wdExportCreateWordBookmarks    # Create a bookmark in the exported document for each Word bookmark

    	
	# start Word
	$application = New-Object "Microsoft.Office.Interop.Word.ApplicationClass" 

	# Make sure inputFile is an absolute path
	$inputFile = Resolve-Path $inputFile
	#$outputFile = Resolve-Path $outputFile
	$outputFile = [System.IO.Path]::ChangeExtension($inputFile, ".pdf")
	
	$application.Visible = $msoFalse
    $document = $application.Documents.Open($inputFile.Path, $msoFalse, $msoFalse, $msoFalse)
    	
	# export presentation to pdf
	$document.ExportAsFixedFormat($outputFile, $wdExportFormatPDF, $msoFalse, $wdExportOptimizeForPrint, $wdExportAllDocument, 1, 1, $wdExportDocumentContent, $false,  $false, $wdExportCreateNoBookmarks)
	
	$document.Close($msoFalse)
	$document = $null
	
	if($application.Windows.Count -eq 0)
	{
		$application.Quit()
	}
	
	$application = $null
	
	# Make sure references to COM objects are released, otherwise word might not close
	# (calling the methods twice is intentional, see https://msdn.microsoft.com/en-us/library/aa679807(office.11).aspx#officeinteroperabilitych2_part2_gc)
	[System.GC]::Collect();
	[System.GC]::WaitForPendingFinalizers();
	[System.GC]::Collect();
	[System.GC]::WaitForPendingFinalizers();

}


# TODO: let directory be given as parameter for the script. If not given, use D:\test
# TODO: improvement: go through subdirectories also and do Word and Powerpoint at the same time !

Get-ChildItem D:\test -File -Filter *pptx -Recurse |
ForEach-Object {
    $output = "Converting " + $_.FullName
    Write-Output $output
    Export-Presentation($_.FullName)
}

Get-ChildItem D:\test -File -Filter *docx -Recurse |
ForEach-Object {
    $output = "Converting " + $_.FullName
    Write-Output $output
    Export-WordDocument($_.FullName)
}

