<#.Synopsis
Returns Summary of a text document.
.DESCRIPTION
Returns the summary of a text document passed to it, depending upon your chosen word limit (Default 100 words). 
.PARAMETER File
Text File with the content to summarize.
.PARAMETER WordLimit
Maximum number of words to be allowed in Summary
.EXAMPLE
PS Root\> Get-Summary -File D:\Document.txt

“Since my letter [of October 28], the FBI investigation team has been working round-the-clock to process and review a large volume of emails from a device obtained in connection with an unrelated criminal investigation,” Mr. Comey said. It was reported that there were 6,50,000 emails on that laptop.

Provide a path to a text file in the cmdlet and it will generate a summary for you, by default it summarizes upto less than or equal to 100 words.
.EXAMPLE
PS Root\> Get-Summary -File D:\Document.txt -WordLimit 150

“Since my letter [of October 28], the FBI investigation team has been working round-the-clock to process and review a large volume of emails from a device obtained in connection with an unrelated criminal investigation,” Mr. Comey said. It was reported that there were 6,50,000 emails on that laptop. Democratic leader Nancy Pelosi said that after yet another exhaustive review of emails to and from Ms. Clinton, the FBI has again given her a clean bill of health. “The FBI’s findings from its criminal investigation of Hillary Clinton’s secret email server were a damning and unprecedented indictment of her judgment. The FBI found evidence Clinton broke the law, that she placed highly classified national security information at risk and repeatedly lied to the American people about her reckless conduct,” he said.

You can also provide a value to '-WordLimit' parameter to increase or decrease the length of summary.
.EXAMPLE
PS Root\> Get-Summary -File D:\Document.txt -Verbose

“Since my letter [of October 28], the FBI investigation team has been working round-the-clock to process and review a large volume of emails from a device obtained in connection with an unrelated criminal investigation,” Mr. Comey said. It was reported that there were 6,50,000 emails on that laptop.
VERBOSE: Content has been summarized from 875 to 48 Words

Mention a '-Verbose' switch to view summarization ratio, i.e, Original number of words to number of words in Summary.
.EXAMPLE
PS Root\> Get-Summary -FromClipBoard -WordLimit 50
Indian Prime Minister Narendra Modi has won the online reader’s poll for TIME Person of the Year, beating out other world leaders, artists and politicians as the most .

Use -FromClipboard switch to summarize the content copied to clipboard

.INPUTS
 None. You cannot pipe objects to Get-Summary.

.LINK
Get-Content
.LINK
http://geekeefy.wordpress.com/

.NOTES

Author  : Prateek Singh 
Twitter : @SinghPrateik
Blog    : http://Geekeefy.wordpress.com

#>
Function Get-Summary
{
[cmdletbinding()]
[Alias('Summary')]
[OutputType([String])]
Param(
        [Parameter(Position=0)] [String] $File,
        [Parameter(Position=1)] [Int] $WordLimit = 100,
        [switch] $FromClipBoard
)

Begin
{
    If($File)
    {
        $Content = Get-Content $File
    }
    elseif($FromClipBoard)
    {
        Add-Type -Assembly PresentationCore
        $Content = [Windows.clipboard]::GetText()
    }
    else
    {
        Write-Host "Please provide a file path or copy content to Clipboard"
    }
}
Process
{
    $TotalWords = 0
    $Summary=@()
    
    #Extracting Best sentences with highest Ranks within the word limit
    $BestSentences = Foreach($Item in (Get-SentenceRank $Content | Sort SentenceScore -Descending))
    {
        #Condition to limit Total word Count
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
    
    If($BestSentences)
    {

        #Constructing a paragraph with sentences in Chronological order
        Foreach($best in (($BestSentences |sort Linenumber).sentence))
        {
            If(-not $Best.trim().endswith("."))
            {
                $Summary += -join ($Best, ".")
            
            }
            else
            {
                $Summary += -join ($Best, "")
            }
        
        }
        
        [String]$Summary

        Write-Verbose "Content has been summarized from $($Content.split(" ").count) to $(([string]$Summary).split(" ").count) Words"
    }
    else
    {
        Write-Warning "Word Limit is too small to summarize the document."
    }
}
End
{

}

 
}

Function Get-Intersection($Sentence1, $Sentence2)
{
    $CommonWords = Compare-Object -ReferenceObject $Sentence1 -DifferenceObject $Sentence2 -IncludeEqual |?{$_.sideindicator -eq '=='} | select Inputobject -ExpandProperty Inputobject
    $CommonWords.Count / ($Sentence1.Count + $Sentence2.Count) /2
}

Function Get-SentenceRank($Content)
{
    $Sentences = $content -split [environment]::NewLine | ?{$_}


    $NoOfSentences = $Sentences.count
    $values = New-Object 'object[,]' $NoOfSentences,$NoOfSentences
    $CommonContentWeight = New-Object double[] $NoOfSentences
    
    #Get important words that where length is greater than 3 to avoid - in, on, of, to, by etc
    $FrequencyDistribution =  $Content.split(" ") |?{-not [String]::IsNullOrEmpty($_)} | %{[Regex]::Replace($_,'[^a-zA-Z0-9]','')} |group |sort count -Descending
    $ImportantWords = $FrequencyDistribution |?{$_.name.length -gt 3} | select @{n='ImportanceWeight';e={$_.Count * 0.01}}, @{n='ImportantWord';e={$_.Name}} -First 10

    Foreach($i in (0..($NoOfSentences-1)))
    {
        $ImportanceWeight = 0

        #Score each Sentence on basis of words common in every other sentence
        #More a sentence has common words from all other sentences, more it defines the complete document
                
        Foreach($j in (0..($NoOfSentences-1)))
        {
            $WordsInReferenceSentence = $Sentences[$i].Split(" ") | Foreach{[Regex]::Replace($_,'[^a-zA-Z0-9]','')}
            $WordsInDifferenceSentence = $Sentences[$j].Split(" ") | Foreach{[Regex]::Replace($_,'[^a-zA-Z0-9]','')}
        
            $CommonContentWeight[$i] = $CommonContentWeight[$i] + (Get-Intersection  $WordsInReferenceSentence $WordsInDifferenceSentence)
        }

        Foreach($Item in $WordsInReferenceSentence |select -unique)
        {
            #Keep adding ImportanceWeight if an Important word found in the sentence
            If($Item -in $ImportantWords.ImportantWord)
            {
                $ImportanceWeight += ($ImportantWords| ?{$_.ImportantWord -eq $Item}).ImportanceWeight
            }
        }
    
        ''| select  @{n='LineNumber';e={$i}},@{n='SentenceScore';e={"{0:N3}"-f ($CommonContentWeight[$i]+$ImportanceWeight)}} ,  @{n='CommonContentScore';e={"{0:N3}"-f $CommonContentWeight[$i]}}, @{n='ImportanceScore';e={$ImportanceWeight}}, @{n='WordCount';e={($Sentences[$i].Split(" ")).count}} , @{n='Sentence';e={$Sentences[$i]}}
    }
}

Export-ModuleMember -Function Get-Summary -Alias "*"
