# CombineDCX
# Reorders DCX files and sorts them into correct order before re-saving as DCX
#
use strict;
use File::Basename;
use File::Path;

my($EvenIsReversed, $EvenPageSuffix, $OddPageSuffix);
my(@FileList, $TempFolder, $TempFileFormat, $MultiPageFormat);

#Global
$EvenIsReversed=1;	  # by default, even numbered pages are in reverse order.
						 # Change this value to False if it is not.
$TempFolder="_temp_dcx_conversion"; 	# creates a _temp folder to save all the extracted files. Once finished,
						 # content of _temp folder is deleted (at the end)
$TempFileFormat=".pcx";  # Use PCX as temp file when the files are extracted and converted from DCX
$EvenPageSuffix = "_e";  # all even page dcx should end with _e. Ex: myText_e.dcx
$OddPageSuffix="_o";	 # all odd page dcx should end with _o. Ex: myText_o.dcx
$MultiPageFormat=".dcx"; # Using DCX. However, we may opt for PDF or other format in the future.

#local
my $currentFileNumber ;;			# starts from 0 and up.
my $oddCounter ;
my $evenCounter  ;
my ($from, $to);
# Begin

mkdir ($TempFolder);	 # create temp folder

my($oddPageSize, $evenPageSize, @tempArray);
my $currentFileName;
my($cmdline);
# Find Odd Paged DCX first , and then the even pages
foreach(<*$OddPageSuffix$MultiPageFormat>) {   # "myPic_o.dcx"

	unlink <$TempFolder/*.*>;	# start with empty temp folder.

	extractDcx("$_");
	$_ = basename($_,$MultiPageFormat);  # remove Extension
	s/$OddPageSuffix//i;  # remove extension and _e,_o.

	# count # of odd pages it generated.
	@tempArray=<$TempFolder/*>;  #get odd page #
	$oddPageSize=@tempArray;
	print "odd page size= $oddPageSize\n";
	@FileList = ($_, @FileList);
	$currentFileName = $_;
	# even Page extract. extractDcx("myPic_e.dcx")
	extractDcx("$_$EvenPageSuffix$MultiPageFormat");
	@tempArray=<$TempFolder/*>;
	$evenPageSize= @tempArray-$oddPageSize;
	print "even page size= $evenPageSize\n";

   # rename all files to PCX since ImageMagick appends Number after extracting them.
   foreach (<$TempFolder/*>) {
	   print ("Renaming $_ to $_$TempFileFormat\n");
	   rename ("$_","$_$TempFileFormat");
   }

   # combine DCX
	$currentFileNumber = "00000";			 # starts from 0 and up.
	$oddCounter = 0;
	$evenCounter = $EvenIsReversed ? $evenPageSize-1: 0 ;
   #my $i, $j;
   for ($oddCounter=0; $oddCounter < $oddPageSize; $oddCounter++) {
		   # add odd sheet
		   $from= "$TempFolder/$currentFileName$OddPageSuffix$TempFileFormat.$oddCounter$TempFileFormat";
		   $to = "$TempFolder/$currentFileNumber$TempFileFormat";
		   print "rename $from => $to\n";
		   rename ($from,$to);
		   $currentFileNumber++;	#Magic increment.

		   #add even sheet . Some doc may end with odd page so we watch for it here.
		   if ($evenCounter>=0 and $evenCounter<$evenPageSize) {
			   $from= "$TempFolder/$currentFileName$EvenPageSuffix$TempFileFormat.$evenCounter$TempFileFormat";
			   $to = "$TempFolder/$currentFileNumber$TempFileFormat";
			   print "rename $from => $to\n";
			   rename ($from,$to);
			   $evenCounter = $EvenIsReversed ? $evenCounter-1: $evenCounter+1 ;
			   $currentFileNumber++;	#Magic increment.
		   }
		   #if ($evenCounter <0 or $evenCounter >= $evenPageSize) {last;}
   }
   # Now, call Convert
	#my($destinationFile) = basename($fileName,$MultiPageFormat);  # remove Extension
	#$destinationFile = "$destinationFile$TempFileFormat";
	$cmdline=  "d:/_app/ImageMagick/convert.exe $TempFolder/*$TempFileFormat $currentFileName$MultiPageFormat"	;
	print "Final Convert : $cmdline\n";
	system( $cmdline);

}

# Final Step
rmtree($TempFolder);


#Call to Extract DCX to PCX, etc.
sub extractDcx {
	my($fileName)=$_[0] ;
	my($destinationFile) = basename($fileName,$MultiPageFormat);  # remove Extension
	$destinationFile = "$destinationFile$TempFileFormat";
	my($cmdline);
	$cmdline=  "d:/_app/ImageMagick/convert.exe $fileName $TempFolder/$destinationFile"  ;
	print "Extracting : $cmdline\n";
	system( $cmdline);
}

# Combine all PCX files into DCX.
sub makeDcx {

}
