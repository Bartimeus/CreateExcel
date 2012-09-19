use Spreadsheet::XLSX;
use Excel::Writer::XLSX;
use Excel::Writer::XLSX::Worksheet;
 
 #открываем на чтение
 my $inExcel = Spreadsheet::XLSX->new ('test.xlsx');
 my $wsRead = $inExcel->worksheet(0);
 
 #создаем новый 
 my $outExcel  = Excel::Writer::XLSX->new('out.xlsx');
 my $wsWrite = $outExcel->add_worksheet();
 
 my $chart = $outExcel->add_chart( type => 'line', embedded => 1 );

 my $i = 0;

 foreach my $sheet (@{$inExcel->{Worksheet}}) 
 {
        $sheet->{MaxRow} ||= $sheet->{MinRow};
      
         foreach my $row ($sheet->{MinRow} .. $sheet->{MaxRow}) 
		 {
                $sheet->{MaxCol} ||= $sheet->{MinCol};
				
				my $j = 0;
			
                foreach my $col ($sheet->{MinCol} ..  $sheet->{MaxCol}) 
				{
                        my $cell = $sheet ->{Cells} [$row] [$col];
						$wsWrite->write($i , $j, $cell->{Val});
						$j++;
                }
				
				$i++;
 
        }
		
		foreach my $col ($sheet->{MinCol} ..  $sheet->{MaxCol})
		{
			$chart->add_series(
			values     => [ 'Sheet1', 0, $sheet->{MaxRow}, $col, $col],
			);
		}		
 }

 $wsWrite->insert_chart('E2', $chart, 100, 10);
 


	
	