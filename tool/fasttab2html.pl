#!/usr/bin/perl 
#=====================================================================
# Convert fasttab text to HTML table. 
#
# My First Perl App
# (c) Dan Kim
# see pod below for more info
#=====================================================================
package Fasttab2html;
use strict;
use warnings;
use 5.004;
use Carp;
use File::Basename;
#use CGI qw(:standard); Don't need big CGI.pm for this one
use The;	# my lib

#VAR
my $_current_column=0;
my $_continuous_newline=1; #keeps track of # consecutive empty lines 
my $_border=1;		#<table border=>
my $_valign='top';	#<tr valign=>
my $_td_style='';		#<td><{style}...></td>
my $_is_block=0;  # prev line has "[" 
my $_previous_line_is_block_begin=0;
#----------------------------------------------------------------------------
#main loop for the app
# arg: array of lines of text
# Returns: array of lines of text after conversion
#----------------------------------------------------------------------------
sub parse_text {
	my (@_original_text)=@_;
	#print @_original_text;
	my @_new_text;
	foreach (@_original_text) {
		push(@_new_text,parse_line($_));
	}
	return @_new_text;
}

#----------------------------------------------------------------------------
# parse each line
# return: new line that has been converted.
#----------------------------------------------------------------------------
sub parse_line {
	$_=shift;
	my $_text="";		# holds our string that will be returned
	my $_column_difference;	
	my $i;			# used in loop
	#my $_is_all_cap=0;	# true if the text is written in CAPS
	#my $_is_in_bracket=0;	# true if text is in []. ie [abc....]
	
	/^\t*/;	# count # of tabs
	my $_num_of_tabs=length($&);
	my $_number_of_tab_within_block=0;
	s/^\t*//;	# strip beginning initial tabs except for those that are in "[]" block.
	
	my $_is_all_cap=(/[a-z]/)?0:1;  # check for ALL-CAPS
	my $_is_in_bracket=(/^\[/ && /\]/); #check if it is in [ ]

	# We have to change col in our table.
	$_column_difference=$_num_of_tabs-$_current_column;
	
	if ($_is_block) { 		# in block mode "[\n\n...]"
		$_column_difference=0;	# current block should be in same level as prev line
		$_number_of_tab_within_block=$_num_of_tabs-$_current_column;
		#warn "number_of_tab=$_number_of_tab_within_block"
	}
	if ($_column_difference >0) {
	#if ($_current_column < $_num_of_tabs) {
		# add columns
		foreach($i=0; $i<$_column_difference;$i++) {
			$_text.=html_td_end().html_td();		# </td><td>
		}
			$_current_column=$_num_of_tabs;
	}
	elsif ($_column_difference <0) {
	#elsif ($_current_column > $_num_of_tabs) {
		# add columns
		#if ($_num_of_tabs==0) {
		
		#Start a new table if prev lines were all blanks
		if ($_continuous_newline>1) {	
			$_text.="</tr></table>".html_table().html_tr();
		}
		# Normal Texts
		else {
			$_text.=html_td_end()."</tr>".html_tr();
		}
		#else {
			foreach($i=0; $i<$_num_of_tabs;$i++) {
				$_text.="<td></td>";
			}
			#$_text.="</tr><tr><td>";
		#}
		$_current_column=$_num_of_tabs;
		
		# a line followed by blank line or All CAP text indicates 
		# it is a heading or a topic. Therefore, make it BOLD.
		if ($_continuous_newline || $_is_all_cap ||$_is_in_bracket) {$_td_style="b";}
		
		$_text.=html_td();	
	}
	else {
		if (! $_continuous_newline) {
			if (!$_previous_line_is_block_begin) { 
				$_text.="<br>";
				$_previous_line_is_block_begin=0;
				}
			for ($i=0;$i<$_number_of_tab_within_block;$i++) {
				$_text.="\t";
			}
		}
	}

	my $t=$_;
	$t=~s/^[ \t]//;
	if ($t eq "\n" || $t eq "\r\n") {
		#warn "continue line detected";
		$_continuous_newline++;
	}
	else {
		$_continuous_newline=0;
	}
	if (/^\[/ && ! /\]/) { 	# begin block "["
	  $_is_block=1;
	  $_text.="<pre>";
	  $_previous_line_is_block_begin=1;
	  #$_current_column=$_num_of_tabs;
	}
	elsif (!/^\[/ && /\]$/) { # end block "]"
	  $_is_block=0;
	  $_text.="</pre>";
	}
	else {						# normal text, not "[" or "]"
		$_text.=$_;
	}
	return $_text;
}



#----------------------------------------------------------------------------
# converts txt to html
# 
#----------------------------------------------------------------------------
sub convert_file {
	# Get filename and save the text in buffer.
	my $filename=shift || die ("no filename.");
	open(INPUT_FILE ,$filename)  || croak ("input file open failed");
	#my @text;

	# make sure that 1st line contains "=table"
	my $first_line=<INPUT_FILE>;	# read 1st line of text
	chomp ($first_line);	
	$first_line=~s/\s+//g;		# remove all whitespace
	$first_line=~tr/A-Z/a-z/;	#lowercase 
	($first_line eq "=table") || croak ("Not FastTable format");
	
	my ($basename,$path,$extension)=fileparse($filename,'\..*');
	#my ($basename,$path,$extension)=fileparse($filename);
	my $output_filename=$path.$basename.".html";
	open(OUTPUT_FILE,">$output_filename") || croak ("output file open failed");
	#warn "filename=$filename, basename=$basename , path=$path output=$output_filename";
	#@text=<INPUT_FILE>;
	print OUTPUT_FILE "<html>\n<header><title>$filename</title></header>\n";
	print OUTPUT_FILE"<body>\n";
	print OUTPUT_FILE html_table(), html_tr(), html_td();
	while (<INPUT_FILE>) {
			print OUTPUT_FILE parse_line($_);
		}
	print OUTPUT_FILE html_td_end(),"</tr></table>\n";
	print OUTPUT_FILE"</body></html>\n";
	close(INPUT_FILE);
	close (OUTPUT_FILE);
}
#print (parse_text(@text));

my $filename=shift || die ("no filename.");
convert_file($filename);

#=====================================================================
# MODULE RETURN VALUE
#=====================================================================
1;
__END__
#=====================================================================
# Module Documentation (POD)
#=====================================================================
=head1 fasttab2html
Convert fasttab to html
=cut
