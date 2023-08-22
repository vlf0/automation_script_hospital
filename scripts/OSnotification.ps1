$Testbut = New-BTButton -Content 'Got it' -Arguments None
$User = 'User'
$Header = New-BTHeader -Id '001' -Title 'ВНИМАНИЕ!'
New-BurntToastNotification -Text "$User, нет файлов excel в папке программы!", 'Добавьте файлы и повторите попытку!' -Header $Header -Sound Alarm4 -Button $testbut;