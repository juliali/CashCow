$IE= new-object -com "InternetExplorer.Application"
$IE.navigate2("http://www.google.com/?tbm=nws&as_qdr=m3")

$n=1
while ($IE.busy) {
sleep -milliseconds 100
}

$actionArray = @();

Get-Content F:\tmp\ActionList.txt | ForEach-Object {

$actionArray += $_.ToString().Trim();
}

Get-Content F:\tmp\CompanyNameList.txt | ForEach-Object {

$List = New-Object Collections.Generic.List[String]

$company = $_.ToString().Trim()

foreach ($action in $actionArray)
{

$query = $company + " " + $action; 

$IE.visible=$true
$IE.Document.getElementById("q").value=$query
$IE.visible=$true

$IE.document.forms | 
    Select -First 1 | 
        % { $_.submit() }

while ($IE.busy) {
sleep -milliseconds 100
}

$element = $IE.document.getElementById("search");

    $itemArray = @();
    
    foreach($item in $element.getElementsByTagName("h3"))
    {
        $itemArray += $item.innerText;
        
    };

    $urlArray = @();
    foreach($url in $element.getElementsByTagName("a"))
    {
        if ($url.className -eq "l _HId" )
        {
            $urlArray +=¡¡$url.href;
        }
    }
    
    $contentArray = @();
    foreach($content in $element.getElementsByTagName("div"))
   {
    
        if ( $content.className -eq "st" ) 
       {            
            $contentArray += $content.innerText;
        }
    
    }



    for($i=0; $i -le 9; $i++)
    {
        $List +=  $company + " | " +¡¡$action + " | " + $itemArray[$i] + " | " + $contentArray[$i] + " | "  + $urlArray[$i]  + "`n" ;
    }
    
    Write-Host $List + "`n";   


$List | out-file -append "F:\tmp\news.csv";

Write-HOST "OUTPUT" $n
$n++
}
}

