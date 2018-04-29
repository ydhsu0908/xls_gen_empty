use Text::Iconv;
my $converter = Text::Iconv -> new ("utf-8", "windows-1251");

# Text::Iconv is not really required.
# This can be any object with the convert method. Or nothing.

use Spreadsheet::XLSX;

open $if, "< ./xls_gen_empty.xlsx"  or die "Can't open file to read\n";


my $excel = Spreadsheet::XLSX -> new ('xls_gen_empty.xlsx', $converter);

foreach my $sheet (@{$excel -> {Worksheet}}) {

   open $of, "> ./$sheet->{Name}.v" or die "Can't open file to write\n";

   $sheet -> {MaxRow} ||= $sheet -> {MinRow};

   foreach my $row ($sheet -> {MinRow} .. $sheet -> {MaxRow}) {
     $sheet -> {MaxCol} ||= $sheet -> {MinCol};
   }
    
   my $min_row = $sheet -> {MinRow};
   print "min_row:".$min_row."\n";
   my $max_row = $sheet -> {MaxRow};
   print "max_row:".$max_row."\n";
   my $min_col = $sheet -> {MinCol};
   print "min_col:".$min_col."\n";
   my $max_col = $sheet -> {MaxCol};
   print "max_col:".$max_col."\n";
   
    printf $of "//Created by xls_gen_empty.pl \n";
    $datestring = localtime();
    printf $of "//$datestring\n\n";
    printf $of "`timescale 1ns/1ps"."\n";
    printf $of "module ".$sheet->{Name}."(\n";
    
    foreach my $i ($min_row+1 .. $max_row) {
        $pin = $sheet -> {Cells} [$i] [$min_col+0]; ##Pin
        $dir = $sheet -> {Cells} [$i] [$min_col+1]; ##Dir
        $des = $sheet -> {Cells} [$i] [$min_col+2]; ##Des
        
        my @dir_arr = split(/ /, $dir-> {Val});
        my @pin_arr = split('\[', $pin-> {Val});
        
        $pin_name = $pin_arr[0] . ";";
        $pin_bus  = $pin_arr[1];
        
        if($pin_arr[1]) {
          $pin_bus  = "[".$pin_arr[1];
        } else {
          $pin_bus  = " ";
        }
        
        
        if($dir_arr[0] eq "IN"){
           printf $of " input  %-10s %-20s //%s\n", $pin_bus, $pin_name, $des-> {Val};
        } elsif ($dir_arr[0] eq "OUT") {
           printf $of " output %-10s %-20s //%s\n", $pin_bus, $pin_name, $des-> {Val};
        } else {
           printf $of " inout  %-10s %-20s //%s\n", $pin_bus, $pin_name, $des-> {Val};
        }
        
    }
    printf $of ");\n";
    printf $of "endmodule \n";
}
