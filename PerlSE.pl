use Win32::OLE::Lite;
BEGIN { Win32::OLE->Initialize(); }

sub _eval_init_
{
	# internal variables
	$_eval_code_ = '';
	$_eval_line_ = 0;
	$_eval_errline_ = 0;
	$_eval_errcode_ = '';
}

sub _eval_
{
	$_eval_code_ = $_[0];
	$_eval_line_ = $_[1];
	eval $_[0] . "\ndefined \&Win32::ASP::_END && \&Win32::ASP::_END;\n";
	if(!$@)
	{	# no error;
		return ;
	}
	else
	{
		my $w = $@;
		$w =~ s/at \(eval (\d+)\) //;
		my $eb=$1;						# eval block
		$w =~ s/ line (\d+)//;
		my $el=$1-1;					# line number in eval block
		$_eval_errline_ = $_eval_line_+$el;
		my @lines = split(/^/m, $_eval_code_);
		if($el > 0 ) {
			$_eval_errcode_ = $lines[$el-1] . $lines[$el];
		} else {
			$_eval_errcode_ = $lines[$el] . $lines[$el+1];
		}
		return $w;
	}
}

