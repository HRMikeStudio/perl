



use Net::SMTP;
#--发邮件包
use MIME::Lite;
#--多媒体附件包
use Net::SMTP::SSL;
#--能用SSL登陆包？


use strict;
my $file;
#my $dir = "./test/";
my $dir = "/Volumes/limited/Hr/Payroll/SALARY/2014/2014 Employee Salary Slip/Value version/";
#--Salary slips storge path.
 
my @mail_to;
#--Define mail addresses object
print "$dir\n";
opendir(DIR,$dir);
@mail_to = readdir DIR;
print @mail_to;


foreach my $file (@mail_to) 
{
    next if $file eq "." or $file eq "..";
    #--文件夹中的"."和".."忽略
    next if !(-f $dir.$file && $file =~ /\.xlsx/i);
   #--如果不是文件夹中的"xlsx"文件，忽略
    
    my $mail = $file;
    print $mail, "\n";
    #$mail =~ s/\.xlsx$//i;
    $mail =~ s/\.xlsx$//i;
    print $mail, "\n";
    
    my $msg = MIME::Lite-> new(
            From    => 'chubow@opera.com',
    #        To      => 'chubow@opera.com',        
            To      => $mail,
            Subject => 'Salary slip of May 2014 ',
              Type    =>'multipart/mixed',
    );
    #--定义邮件对象，发件箱，收件箱，邮件主题
    $msg->attach(
        Type => 'TEXT',
        Data => "HI, Please check your salary slip. Any question just let me know. \n\n\n This mail is sent out by Salary slip program automatically. Powered by chubow.",
        );
    #--邮件正文
    $msg->attach(
#       Type     => 'application/msword',
#        Type     => 'multipart/mixed',
        
        Type     => 'AUTO',
        #--附件的类型。
        Path     => $dir.$file,
#        Path     => '/Volumes/limited/Hr/Payroll/SALARY/2014/2014 Employee Salary Slip/Value version',
#        Filename => 'test.xlsx',
#        Filename => $file,
#        Disposition => 'attachment',
    );
    #--附件的路径. 
print "mail:$mail, file:$file, msg:$msg\n";



    my $smtp = new Net::SMTP::SSL(  
        'smtp.gmail.com',  
        Port    =>      465,  
    #    User    =>      'dabing289',  
    #    Password=>      '722556123',
    );
    $smtp->auth( 'chubow@opera.com','YUiMMgK6');  
    $smtp->mail('chubow@opera.com');  
    #sigle receipt
    #$smtp->to('chubow@opera.com');  
    #multi-receipt
    $smtp->to($mail); 
    #--传入收件人地址
    $smtp->data();  
    my $mail_content = $msg->as_string;
    #print "mail_content:$mail_content\n";
    print "123\n";
    $smtp->datasend($mail_content);
    #--传入邮件正文  
    $smtp->dataend();  
    $smtp->quit;  
    sleep 3;
    
    
}
