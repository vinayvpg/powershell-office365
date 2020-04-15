    Connect-PnPOnline -Scopes "User.Read","User.ReadBasic.All","Group.Read.All"

    $accessToken = Get-PnPAccessToken

    $global:groupId = "86d665c9-955f-4017-83a5-1807362d0b65"

    function GetPost($topic, $topicThreadId) {
        $postsEndPoint = "https://graph.microsoft.com/v1.0/groups/" + $global:groupId + "/threads/" + $topicThreadId + "/posts?`$top=1000"

        $postsResp = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken"} -ContentType "application/json" -Uri $postsEndPoint -Method Get

        if($postsResp -ne $null) {
            $posts = $postsResp.value
            for($i=0; $i -lt $posts.Length; $i++) {
                $postId = $posts[$i].id
                $postDate = $posts[$i].createdDateTime
                $sender = $posts[$i].from.emailAddress.name
                $commentHtml = $posts[$i].body.content
                
                $htmlParser = New-Object -Com "Htmlfile"
                $htmlParser.IHTMLDocument2_write($commentHtml)
                $comment = $htmlParser.getElementById("jSanity_hideInPlanner").parentNode | % innerText | % {$_ -replace ","," "} | % { $_ -replace "`n", " "} |  % { $_ -replace "`r", " "}

                Write-Host "---->Processing post Id: $postId by sender: $sender"
                $comment

                "$topic `t $topicThreadId `t $sender `t $postDate `t $postId `t $comment" | out-file $global:reportPath -Append 
            }
        }
    }

    $endPoint = "https://graph.microsoft.com/v1.0/groups/" + $global:groupId + "/threads?`$top=1000"

    if(![string]::IsNullOrWhiteSpace($endPoint)) {
        #$req = [System.Net.WebRequest]::Create($endpoint)
        #$req.Accept = "application/json;odata=verbose"
        #$req.Method = "GET"
        $resp = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken"} -ContentType "application/json" -Uri $endPoint -Method Get

        if($resp -ne $null) {
            write-host "Start - $(Get-Date)`n" -ForegroundColor Yellow

            $timestamp = Get-Date -Format s | % { $_ -replace ":", "-" }

            $currentDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
            $reportCSVPath = $currentDir + "\planner-comments-$global:groupId" + $timestamp + ".csv"
            Write-Host "You did not specify a path to the report csv file. The report will be created at '$reportCSVPath'" -ForegroundColor Cyan

            $global:reportPath = $reportCSVPath

            "Topic `t TopicThreadId `t Sender `t CommentDateTime `t CommentPostId `t Comment" | out-file $global:reportPath

            #[System.Net.WebResponse] $resp = $req.GetResponse()
            #[System.IO.Stream] $respStream = $resp.GetResponseStream()

            #$readStream = New-Object System.IO.StreamReader $respStream
            
            #$ret = $readStream.ReadToEnd() | ConvertFrom-Json

            $ret = $resp.value

            if($ret -ne $null) {
                if($ret.Count -ne 0) {
                    for($i=0; $i -lt $ret.Length; $i++) {
                        $topic = $ret[$i].topic
                        $topicThreadId = $ret[$i].id

                        Write-Host "--------------------------------------------------------------"
                        Write-Host "Topic: $topic Thread Id: $threadId"

                        GetPost $topic $topicThreadId
                    }
                }
            }
        }
    }
