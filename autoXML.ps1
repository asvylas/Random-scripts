[xml]$xmlData = get-content C:\Users\svylasan\Desktop\data.xml
[xml]$xmlMastersheet = get-content C:\Users\svylasan\Desktop\mastersheet.xml
    
    for($i=0; $i-lt $xmlData.INSTAT.Envelope.Declaration.Item.CN8.CN8Code.Length; $i++)
    {
        for ($z=0; $z-lt $xmlMastersheet.root.row.Length; $z++)
        {
            if($xmlMastersheet.root.row.codenumber[$z] -eq $xmlData.INSTAT.Envelope.Declaration.Item.CN8.CN8Code[$i])
            {
            $xmlData.INSTAT.Envelope.Declaration.Item[$i].goodsDescription = $xmlMastersheet.root.row.desc[$z]
            echo $xmlData.INSTAT.Envelope.Declaration.Item[$i].goodsDescription
            }
        }

    }
    
$xmlData.save("C:\Users\svylasan\Desktop\File.xml")
echo ""
echo ""
echo "TASK COMPLETED!"
echo ""
echo ""