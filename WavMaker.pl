#!/usr/bin/perl
#

use strict;
use warnings;
use Config;
use File::Find;

# Create Hash for storing the main menu
my %options = (
	"1" => {"Subroutine" => \&SingleWav, "Read" => "Single Wav File"},
	"2" => {"Subroutine" => \&ExcelParse, "Read" => "Parse Excel (.xls) document for Prompts."},
	"3" => {"Subroutine" => \&SetDir, "Read" => "Change Output Directory"},
	"4" => {"Subroutine" =>	\&Exit, "Read" => "Exit Program"},
);
my $key = keys %options;


# Initialize global variables
my $input = my $dirname = my $soxPath = "";

START:

	#&FindSOX;

	print "\nThe following voices are installed: \n";

	#------------------------------------------------
	# Run test procedure
	#------------------------------------------------

	&SAPItest;

	#------------------------------------------------
	# Display all installed voices 
	#------------------------------------------------
	my %voicelist = &SAPIgetVoices;
	my $a = '';
	print "\n";
	foreach $a (sort keys %voicelist)
	{
		print "$voicelist{$a} = $a\n";
	}

	print "Who would you like to record in? \n";
	chomp(my $voice = <STDIN>);

	&SetDir;

MAINMENU:

	print "\n";
		
	# The Main Menu
	print "What would you like to do? \n";
	# Print out each option in the options hash
	foreach $key (sort keys %options)
	{
		print $key, ": ", $options{$key}{"Read"}, "\n";
	}
	# User input
	chomp($input = <STDIN>);	
	# Get the subroutine Name
	my $option = $options{$input}{"Subroutine"};
	print "\n";

	# Invoke subroutine stored in variable $option
	$option -> ();

goto MAINMENU;

sub LoadParams
{
	#####################################################################
	# LoadParams
	#-------------------------------------------------------------------#
	# Used to load in optional commmand-line params for other use case  
	# Example
	#	-s "My text I want to speak" "out.wav"
	#	-x "path/to/my/spreadsheet"
	#####################################################################
	
	if ($ARGV[0] == undef)
	{
		goto START;
	}
	elsif ($ARGV[0] == "-s")
	{
		use Cwd;
		my ($text, $outFile) = $ARGV[1], getcwd.$ARGV[2];
		
		&SAPIwave($text, $voice, "$outFile._temp");

		# Convert audio using sox.  	
		&ConvertFile("", $file);
	}
	elsif ($ARGV[0] == "-x")
	{
		use Cwd;
		
	}
}

sub SetDir
{
	#####################################################################
	# SetDir
	#-------------------------------------------------------------------#
	# Simply sets/changes the output folder
	#####################################################################
	
	print "Where would you like the file to be written? \n";
	chomp($dirname = <STDIN>);
	$dirname =~ s/\//\\/g;
}

sub ExcelParse
{
	use Spreadsheet::ParseExcel;
	
	#####################################################################
	# ExcelParse
	#-------------------------------------------------------------------#
	# Parses Excel Doc for audio prompts and file names
	# Column A is audio file names 
	# Column B is text to read out
	#####################################################################
	
	print "Enter the path to the Excel doc containing the audio prompts \nColumn A is the name of the file, Column B is the text.\n";
	my $parser = Spreadsheet::ParseExcel->new;
	chomp(my $excelPath = <STDIN>);
	$excelPath =~ s/\//\\/g;

	my $workbook  = $parser->parse($excelPath) or die;
	my $worksheet = $workbook->worksheet(0);

	my %data;
	for my $row ( 0 .. $worksheet->row_range ) 
	{
		my $file = $worksheet->get_cell( $row, 0 )->value;
		
		# add .wav if needed
		if (!($file =~ /.wav$/))
		{
			$file = $file.".wav";
		}
		
		my $text = $worksheet->get_cell( $row, 1 )->value;
		
		# use the name of the file if no value given
		if ($text eq "")
		{
			$text = $file;
		}
    
		&SAPIwave($text, $voice, "$dirname\/$file._temp");

		# Convert audio using sox.  
		&ConvertFile($dirname, $file);
	}
}

sub SingleWav
{
	######################################################################
	# SingleWav
	#--------------------------------------------------------------------#
	# Used to do just one wav file
	######################################################################

	print "What would you like me to say? \n";
	chomp(my $text = <STDIN>);

	print "And What do you want the file to be called? \n";
	chomp(my $file = <STDIN>);

	&SAPIwave($text, $voice, "$dirname\/$file._temp");

	# Convert audio using sox.  	
	&ConvertFile($dirname, $file);

}

sub FindSOX
{
	######################################################################
	# FindSOX
	#--------------------------------------------------------------------#
	# Figure out where sox.exe is so we can use it to convert files
	######################################################################
	
	my $results = my $soxPath = my $path = "";
	my @programDirectories = ("D:\\Program Files (x86)\\", "D:\\Program Files\\", "E:\\Program Files (x86)\\", "E:\\Program Files\\", "C:\\Program Files (x86)\\", "C:\\Program Files\\");
	
	#Find all the directories under @programDirectories that start with "sox"
	
	foreach $path (@programDirectories)
	{
		#SoX has been found
		if ($results ne "")
		{
			last;
		}
		
		#If directory exists
		if (-d $path)
		{
			opendir( my $DIR, $path );
			while ( my $entry = readdir $DIR ) 
			{
				if ($entry =~ /^sox/)
				{
					$soxPath = $path.$entry."\\sox.exe";
					if (-f $soxPath)
					{
						print "SOX Found at: $soxPath\n";
						$results = "\"".$soxPath."\"";
						last;
					}
				}
			}
		}
	}
	return $results;
}

sub ConvertFile
{
	my ($directory, $file) = @_;
	
	#Find where SoX is if necissary
	if ($soxPath eq "")
	{
		$soxPath = &FindSOX;
	}
	
	#SoX not installed
	if ($soxPath eq "")
	{
		print "SOX not found under Program Files (x86)\\ or Program Files\\ in the C, D, or E drives\nMake sure you have SOX installed and try again.\n";
		my $resp = <STDIN>;
		die;
	}
	
	print "$soxPath -V \"$directory\\$file._temp\" -r 8000 -b 8 -c 1 -e u-law \"$directory\\$file\"\n";
	system("$soxPath -V \"$directory\\$file._temp\" -r 8000 -b 8 -c 1 -e u-law \"$directory\\$file\"");
	system("del \"$directory\\$file._temp\"");
}

sub SAPItest
{
	#####################################################################
	# SAPItest
	#-------------------------------------------------------------------#
	# Speaks each voice installed in the SAPI environment.
	#####################################################################
	

	my $tts = Win32::OLE->new("Sapi.SpVoice") or die "Sapi.SpVoice failed";

	for(my $VoiceCnt=0; $VoiceCnt < $tts->GetVoices->Count(); $VoiceCnt++)
	{
		$tts->{Voice} = $tts->GetVoices->Item($VoiceCnt);
		my $desc = $tts->GetVoices->Item($VoiceCnt)->GetDescription;
		
		my $text = "This is $desc, voice number $VoiceCnt";
		print "[ $text ]\n";
		$tts->Speak("$text", 1);

		$tts->WaitUntilDone(-1);
	}
}

sub SAPIgetVoices
{
	use Win32::OLE;
	#####################################################################
	# SAPIgetVoices
	#-------------------------------------------------------------------#
	# Returns all SAPI voices via a hash
	#####################################################################

	my $tts = Win32::OLE->new("Sapi.SpVoice") or die "Sapi.SpVoice failed";
	my %VOICES;
	for(my $VoiceCnt=0;$VoiceCnt < $tts->GetVoices->Count();$VoiceCnt++)
	{
		my $desc = $tts->GetVoices->Item($VoiceCnt)->GetDescription;

		$VOICES{"$desc"} = $VoiceCnt;
	}
	return %VOICES;
}

sub SAPIwave
{
	#####################################################################
	# SAPIwave
	#-------------------------------------------------------------------#
	# Creates a wave file, worksjust like SAPItalk
	#####################################################################

	my ($text,$voice,$wave) = @_;
	use Win32::OLE;

	my $type=Win32::OLE->new("SAPI.SpAudioFormat");

	# stereo = add 1
	# 16-bit = add 2
	# 8KHz = 4
	# 11KHz = 8
	# 12KHz = 12
	# 16KHz = 16
	# 22KHz = 20
	# 24KHz = 24
	# 32KHz = 28
	# 44KHz = 32
	# 48KHz = 36
	$type->{Type}=4;

	my $stream=Win32::OLE->new("SAPI.SpFileStream");
	$stream->{Format}=$type;

	$stream->Open("$wave",3,undef);

	my $tts=Win32::OLE->new("SAPI.SpVoice");
	$tts->{AudioOutputStream}=$stream;
	$tts->Speak($text,1);

	$tts->WaitUntilDone(-1); 
	$stream->Close(  );
}

sub Exit
{
	exit 0;
}