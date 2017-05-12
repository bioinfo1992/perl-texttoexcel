#!/usr/bin/perl;
#use warnings;
#use strict;
use Cwd;
use Excel::Writer::XLSX;
use File::Basename;
use Encode;

################################################################
#                 Author:chen hao                              #
#                 Date:2016.5.3                                #
#                 Version:1.0                                  #
################################################################


#Main
&excel_export('C:\Users\bio\Desktop\demo.txt', 'C:\Users\bio\Desktop');



=pod
Note:

This function aim to transform text document to Excel2007+ file

Usage:

$parameter1:text file name (include path)
    eg:/home/my/doc/demo.text

$parameter2:export path
    eg:/home/my/output/

$parameter3:export name for excel
    eg:demo --> output :demo.xlsx(defult name eq filename)

$paramater4:table name rule
    eg:demo --> sheet name:demo1,demo2 .etc(defult Sheet1,Sheet2...)

=cut

sub excel_export(){
    if (@_ < 2) {
        #error $parameter
        warn "Missing argument!\n";
        return "false";
    }
    my $currentDIR = getcwd();
    #my $newDIR;
    #Get information
    my ($file, $output, $fileName, $tableName);
    ($file, $output) = @_ if @_ == 2;
    ($file, $output, $fileName) = @_ if @_ == 3;
    ($file, $output, $fileName, $tableName) = @_ if @_ == 4;
    #print "$file, $output, $fileName, $tableName, \n";
    #$parameter
    if (!defined $fileName) {
        $fileName = basename($file, ".txt");
        #$newDIR = dirname($file);
    }
    
    if (!defined $tableName) {
        $tableName = "Sheet";
    }
    #parameter check
    if (!-e $file) {
        warn "No such file in this directory: $file.\n";
        return "false";
    }
    
    if (!-d $output) {
        warn "No such directory : $output \n";
        return "false";
    }
    #chdir $newDIR;
    chdir $output;
    
    open my $in, "<", $file or die "Can't open this file: $!\n";
    # Create a new Excel workbook
    my $workbook  = Excel::Writer::XLSX->new("${fileName}.xlsx");
    
    #Add a worksheet
    my $worksheet;
    if ($tableName =~ /[u4e00-u9fa5]/) {
        $worksheet = $workbook->add_worksheet(decode("gb2312", $tableName));
    }
    else{
        $worksheet = $workbook->add_worksheet($tableName);
    }
    
    #add format
    my $format = $workbook->add_format( num_format => '@' );
    $format->set_align( 'center' );
    
    $worksheet->freeze_panes( 1, 0 );
    
    #write excel
    my $j = 0;
    while (<$in>) {
        chomp;
        my @tmp = split /\t/, $_;
        foreach my $i (0 .. @tmp - 1)
        {
            $tmp[$i] =~ s/^"(.*)"$/$1/;
            $tmp[$i] =~ s/^\s+(.*?)\s+$/$1/;
            if ($tmp[$i] =~ /\d+|\d+\.\d+/) {
                $worksheet->write( $j, $i, $tmp[$i] );
            }
            else{
                if ($tmp[$i] =~ /[u4e00-u9fa5]/) {
                    $worksheet->write_string( $j, $i, decode("gb2312", $tmp[$i]), $format );
                }
                else
                {
                    $worksheet->write_string( $j, $i, $tmp[$i], $format );
                }
            }
        }
        $j++;
    }
    $workbook->close();
    close($in);
    chdir $currentDIR;
    return "true";
}