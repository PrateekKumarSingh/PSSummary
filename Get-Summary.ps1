Function Split-ContentToSentence ($Content)
{
    #([string]$Content -split '. ',0,"simplematch").Trim() | ?{-not [string]::IsNullOrWhiteSpace($_)}  
    $content -split [environment]::NewLine | ?{$_}
}

Function Get-FrequencyDistribution ($Content)
{
    $Content.split(" ") |?{-not [String]::IsNullOrEmpty($_)} | %{[Regex]::Replace($_,'[^a-zA-Z0-9]','')} |group |sort count -Descending
}

Function Get-Intersection($Sentence1, $Sentence2)
{
    $CommonWords = Compare-Object -ReferenceObject $Sentence1 -DifferenceObject $Sentence2 -IncludeEqual |?{$_.sideindicator -eq '=='} | select Inputobject -ExpandProperty Inputobject

    $CommonWords.Count / ($Sentence1.Count + $Sentence2.Count) /2
}

Function Get-SentenceRank($Content)
{
    $Sentences = Split-ContentToSentence $Content
    $NoOfSentences = $Sentences.count
    $values = New-Object 'object[,]' $NoOfSentences,$NoOfSentences
    $SentenceScore = New-Object double[] $NoOfSentences
    
    
    #Get important words that where length is greater than 3 to avoid - in, on, of, to, by etc
    $ImportantWords = Get-FrequencyDistribution $Content |?{$_.name.length -gt 3} | select @{n='Weight';e={$_.Count * 0.01}}, @{n='ImportantWord';e={$_.Name}} -First 10
    $weight = 0

    Foreach($i in (0..($NoOfSentences-1)))
    {
        $NoOfImportantwordsInSentence = 0
                
        Foreach($j in (0..($NoOfSentences-1)))
        {
            $WordsInReferenceSentence = $Sentences[$i].Split(" ") | Foreach{[Regex]::Replace($_,'[^a-zA-Z0-9]','')}
            $WordsInDifferenceSentence = $Sentences[$j].Split(" ") | Foreach{[Regex]::Replace($_,'[^a-zA-Z0-9]','')}
        
            $SentenceScore[$i] = $SentenceScore[$i] + (Get-Intersection  $WordsInReferenceSentence $WordsInDifferenceSentence)
        }

        Foreach($Item in $WordsInReferenceSentence |select -unique)
        {
            #Keep adding weight if an Important word found in the sentence
            If($Item -in $ImportantWords.ImportantWord)
            {
                $weight += ($ImportantWords| ?{$_.ImportantWord -eq $Item}).weight
            }
        }
    
        ''| select  @{n='LineNumber';e={$i}}, @{n='Score';e={"{0:N3}"-f $SentenceScore[$i]}}, @{n='Weight';e={$weight}}, @{n='WordCount';e={($Sentences[$i].Split(" ")).count}} , @{n='Sentence';e={$Sentences[$i]}}
    }
}

Function Get-Summary($Content, $WordLimit)
{

    $TotalWords = 0
    $Summary=@()

    #Extracting Best sentences with highest Ranks within the word limit
    $BestSentences = Foreach($Item in (Get-SentenceRank $Content | Sort Score -Descending))
    {
        $TotalWords += $Item.WordCount

        If($TotalWords -gt $WordLimit)
        {
            break
        }
        else
        {
            $Item
        }
   }
   
    #Constructing a paragraph with sentences in Chronological order
    Foreach($best in (($BestSentences |sort Linenumber).sentence))
    {
        If(-not $Best.endswith("."))
        {
            $Summary += -join ($Best, ".")
        
        }
        else
        {
            $Summary += -join ($Best, "")
        }

    }

    [String]$Summary
}
