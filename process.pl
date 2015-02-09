#!C:\Perl\bin

use Spreadsheet::WriteExcel;
use IO::All;
use Data::Dumper;

my @names = ("Amathria", "Aquendana", "Bankergurl", "Deathtrolla", "Magitroll", "Natilya", "Pacita", "Xarkal");
# Create a new Excel workbook
my $workbook = Spreadsheet::WriteExcel->new("char_inventory.xls");
my $gbank = 0;
my @gboutput = ();

foreach $name (@names){
	my $filename = "C:\\Users\\Public\\Games\\World of Warcraft\\WTF\\Account\\HILLARYP3\\Shadow Council\\$name\\SavedVariables\\GuildLaunchProfiler.lua";
	
	warn Dumper($filename);
	my $in_file = io($filename);
	#warn Dumper($in_file);
	my @lines = io($in_file)->slurp;

	# Add a worksheet
	$worksheet = $workbook->add_worksheet($name);

	#warn Dumper(@lines);
	my $location;
	my $location2;
	my $item;
	my $item_name;
	my $junk;
	my $junk2;
	my $temp;
	my $count;
	my @output = ();

	push(@output, ["ITEM", "COUNT", "LOCATION 1", "LOCATION 2"]);

	foreach $item (@lines){
		if($item =~ /inventory/){
			$location = "INVENTORY";
		}
		if($item =~ /Bag[0123456789]/ || $item =~ /backpack/ || $item =~ /Tab[0123456789]/){
			($junk, $location2, $junk2) = split(/\"/, $item);
		}
		if($item =~ /talents/ || $item =~ /reputation/){
			$location = "";
		}
		if($item =~ /count/ || $item =~ /qty/){
			($junk, $count) = split(/=/, $item);
			$count =~ s/^\s+//;
	  	$count =~ s/\s+$//;
	  	chop $count;
		}
		if($item =~ /bank/){
			($junk, $temp, $junk2) = split(/\"/, $item);
			if($temp eq "bank"){
				$location = "BANK";
				$location2 = "";
			}
		}
		if($item =~ /bag_id/){
			$location2 = "";
		}
		#if($item =~ /contents/ && $location2 eq "" && $location eq "BANK"){
		if($item =~ /guildbank/ && $gbank == 0){
			warn Dumper("Getting Here");
			$worksheet2 = $workbook->add_worksheet('Guild Bank');
			
			push(@gboutput, ["ITEM", "COUNT", "LOCATION 1", "LOCATION 2"]);
			$location = "GUILD BANK";
			
		} elsif($item =~ /guildbank/ && $gbank == 1){
			warn Dumper("Getting Here2");
			$location = "";
		}
	
		if($item =~ /item_name/){
			($junk, $temp, $junk2) = split(/=/, $item);
			($junk, $item_name, $junk2) = split(/\"/, $temp);
			if ($location ne "" && $location ne "GUILD BANK"){
				warn Dumper("Getting Here3");
				#warn Dumper($location2);
		
				push(@output, [$item_name, $count, $location, $location2]);
			} elsif ($location ne "" && $location eq "GUILD BANK"){
				warn Dumper("Getting Here4");
				#$gbank = 1;
				#warn Dumper($location2);
		
				push(@gboutput, [$item_name, $count, $location, $location2]);
			#warn Dumper($item_name);
		}
	}

	
}
$worksheet->write_col('A1', \@output);
	if($#gboutput > 0){
		warn Dumper(@gboutput);
		$worksheet2->write_col('A1', \@gboutput);
	  $gbank = 1;
		@gboutput = ();
	}
}
