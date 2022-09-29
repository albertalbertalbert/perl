use strict;
use Spreadsheet::XLSX;

my ( $inp_xls, $out );
my @values;

if ( !$ARGV[0] ) {
    print "Input file must be specified as first parameter.\n";
    exit;
}
if ( !$ARGV[1] ) {
    print "Output file must be specified as second parameter.\n";
    exit;
}

$inp_xls = Spreadsheet::XLSX->new( $ARGV[0] );
if ( !$inp_xls ) {
    print "Failed to open input spreadsheet.\n";
    exit;
}
my $sheet = $inp_xls->worksheet(0);

process_file();

sub process_file {
    my ( $row_l, $row_h ) = $sheet->row_range();
    my ( $col_l, $col_h ) = $sheet->col_range();
    my $first_row = 1;
    my $first_col = 1;
    my $this_col  = 0;
    for ( my $row = $row_l ; $row <= $row_h ; $row++ ) {

        if ($first_row) {
            $first_row = 0;
            for ( my $col = $col_l ; $col <= $col_h ; $col++ ) {
                if ($first_col) {
                    $first_col = 0;
                }
                else {
                    push @values,
                      [ $sheet->get_cell( $row, $col )->value(), [] ];
                    ##Adds description as first element of array,with a reference to an empty array
                    ##as the second
                }
            }

        }
        else {
            $first_col = 1;
            for ( my $col = $col_l ; $col <= $col_h ; $col++ ) {
                if ( !$first_col ) {
                    push @{ $values[ $col - 1 ][1] },
                      $sheet->get_cell( $row, $col )->value();
                }
                $first_col = 0;

            }

        }

    }
    foreach (@values) {
        print $_->[0] . ":\t";
        my @list = sort { $b <=> $a } @{ $_->[1] };
        for ( my $i = 0 ; $i <= 2 ; $i++ ) {
            print $list[$i] . "\t";
        }
        print "\n";

    }
}
