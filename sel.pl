#!/usr/bin/perl -w
#by Hoang Le P
#6/2/14 4:02PM
#create selection file for update SW or update license for a WRAN network
# how to use: put sel.pl and excel file to a same folder and run script ./sel.pl
# output is save to folder ./output

use Term::ANSIColor qw(:constants);

#use strict;
my $date_time=get_date_time();
my $HOME =  $ENV{"HOME"};
use Spreadsheet::ParseExcel;
my $oExcel = new Spreadsheet::ParseExcel;

#1.1 Normal Excel97

chdir("$HOME/tool_script/selection_creator/");

my $inputfile = get_input_file();

my $oBook = $oExcel->Parse($inputfile);
my ($iR, $iC, $sheet, $cell);



    
$sheet = $oBook->{Worksheet}[0];   

my $batchsize = 30,$batchid=1, $tmp1 = ""; 
$cell = $sheet->{Cells}[1][0]; #store tmp rnc
$tmp1 = $cell->{Val};

print "tmp: ";print $tmp1."\n";

my $ri = 0;

if(-d "./output"){chdir("output");mkdir($date_time);chdir($date_time);}
else{mkdir("output");chdir("output");mkdir($date_time);chdir($date_time);}

for ($iR = $sheet->{MinRow}+1 ; $iR < $sheet->{MaxRow}+1 ; $iR++){ #remove first line
  $cell = $sheet->{Cells}[$iR][0];$rnc = $cell->{Val};
  if ($rnc ne $tmp1) {$tmp1=$rnc;$batchid=1;$ri=0;}

	$ri++;
	
	print "BATCH$batchid ";
	$cell = $sheet->{Cells}[$iR][0];$rnc = $cell->{Val};print "$rnc ";       
	$cell = $sheet->{Cells}[$iR][1];$rbs = $cell->{Val};print "$rbs\n";
	open (MYFILE, ">>$rnc-batch$batchid.sel");	
	
	print MYFILE "SubNetwork=ONRM_ROOT_MO,SubNetwork=$rnc,MeContext=$rbs,ManagedElement=1\n";
	close (MYFILE);
	
	if($ri == $batchsize){$batchid++;$ri = 0}
  

}#end for loop

print "\n\n\>\> Output saved to: "; 
print GREEN, "$HOME/tool_script/selection_creator/output/$date_time\n", RESET;


sub get_date_time{
my $date_time="";
my ($sec,$min,$hour,$mday,$mon,$year,$wday,$yday,$isdst) = localtime(time);
$date_time = sprintf ("%04d%02d%02d_%02d%02d%02d",$year+1900,$mon+1,$mday,$hour,$min,$sec);
#print "$date_time\n";
return $date_time;
}


#automatic get input file name to script
sub get_input_file{
	my $inputfile = "";	
	system("ls | grep -i xls > tmp.txt");
	open (MYFILE, "tmp.txt");
	while (<MYFILE>) 
	{	
		$inputfile = substr($_, 0, -1);  #remove last charactor		
	}
	close (MYFILE); 
	system("rm tmp.txt"); #remove tmp file
	return $inputfile;
}
