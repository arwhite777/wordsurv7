Imports System.Windows.Forms

Public Class ExceptionForm



    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Application.Exit()
        Me.Close()
    End Sub

    Private Sub ExceptionForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim quotes As New List(Of String)
        quotes.Add("Programming today is a race between software engineers striving to build bigger and better idiot-proof programs, and the universe trying to build bigger and better idiots. So far, the universe is winning. -Rick Cook")
        quotes.Add("On two occasions I have been asked, – ""Pray, Mr. Babbage, if you put into the machine wrong figures, will the right answers come out?"" In one case a member of the Upper, and in the other a member of the Lower House put this question. I am not able rightly to apprehend the kind of confusion of ideas that could provoke such a question. -Charles Babbage")
        quotes.Add("Computers are man's attempt at designing a cat: it does whatever it wants, whenever it wants, and rarely ever at the right time. -EMCIC, Keenspot Elf Life Forum")
        quotes.Add("If you lie to the computer, it will get you. -Perry Farrar")
        quotes.Add("No matter how slick the demo is in rehearsal, when you do it in front of a live audience, the probability of a flawless presentation is inversely proportional to the number of people watching, raised to the power of the amount of money involved. -Mark Gibbs")
        quotes.Add("Most software today is very much like an Egyptian pyramid with millions of bricks piled on top of each other, with no structural integrity, but just done by brute force and thousands of slaves. -Alan Kay")
        quotes.Add("Computers are good at following instructions, but not at reading your mind. - Donald Knuth")
        quotes.Add("There are two ways to write error-free programs; only the third one works. -Alan Perlis")
        quotes.Add("Software and cathedrals are much the same – first we build them, then we pray. -Sam Redwine")
        quotes.Add("Why bother with subroutines when you can type fast? -Vaughn Rokosz")
        quotes.Add("Don't get suckered in by the comments … they can be terribly misleading. -Dave Storer")
        quotes.Add("Zawinski's Law: Every program attempts to expand until it can read mail. Those programs which cannot so expand are replaced by ones which can. -Jamie Zawinski")
        quotes.Add("Any sufficiently advanced magic is indistinguishable from a rigged demonstration.")
        quotes.Add("Troutman's First Programming Postulate: If a test installation functions perfectly, all subsequent systems will malfunction.")
        quotes.Add("Troutman's Second Programming Postulate: The most harmful error will not be discovered until a program has been in production for at least six months.")
        quotes.Add("Troutman's Fifth Programming Postulate: If the input editor has been designed to reject all bad input, an ingenious idiot will discover a method to get bad data past it.")
        quotes.Add("Weinberg's Second Law: If builders built buildings the way programmers wrote programs, then the first woodpecker that came along would destroy civilization.")
        quotes.Add("Your program is sick! Shoot it and put it out of its memory.")
        quotes.Add("Much to the surprise of the builders of the first digital computers, programs written for them usually did not work. -Rodney Brooks")
        quotes.Add("bug, n: An elusive creature living in a program that makes it incorrect. The activity of ""debugging"", or removing bugs from a program, ends when people get tired of doing it, not when the bugs are removed. -Datamation")
        quotes.Add("If debugging is the process of removing bugs, then programming must be the process of putting them in. -Edsger W. Dijkstra")
        quotes.Add("Everyone knows that debugging is twice as hard as writing a program in the first place. So if you're as clever as you can be when you write it, how will you ever debug it? -Brian Kernighan")
        quotes.Add("A documented bug is not a bug; it is a feature. -James P. MacLennan")
        quotes.Add("As soon as we started programming, we found to our surprise that it wasn't as easy to get programs right as we had thought. Debugging had to be discovered. I can remember the exact instant when I realized that a large part of my life from then on was going to be spent in finding mistakes in my own programs. -Maurice Wilkes discovers debugging, 1949")
        quotes.Add("The cheapest, fastest and most reliable components of a computer system are those that aren't there. -Gordon Bell")
        quotes.Add("As the trials of life continue to take their toll, remember that there is always a future in Computer Maintenance. -National Lampoon")
        quotes.Add("For a long time it puzzled me how something so expensive, so leading edge, could be so useless.  And then it occurred to me that a computer is a stupid machine with the ability to do incredibly smart things, while computer programmers are smart people with the ability to do incredibly stupid things.  They are, in short, a perfect match. -Bill Bryson")
        quotes.Add("Don’t worry if it doesn’t work right.  If everything did, you’d be out of a job. -Mosher’s Law of Software Engineering")
        quotes.Add("If McDonalds were run like a software company, one out of every hundred Big Macs would give you food poisoning, and the response would be, ""We're sorry, here’s a coupon for two more."" -Mark Minasi")
        quotes.Add("A computer lets you make more mistakes faster than any invention in human history–with the possible exceptions of handguns and tequila. -Mitch Radcliffe")
        quotes.Add("I like what you're doing.  If it worked it'd be great. -Zach Yoder")
        Dim rnd As New Random
        Me.lblQuote.Text = quotes(rnd.Next(0, quotes.Count))
        '                             _                                                 
        '                            | \                                             Oh great, what did you do this time?   
        '                           _|  \______________________________________     /    
        '                          - ______        ________________          \_`,       
        '                        -(_______            -=    -=                   )      
        '                                 `--------=============----------------`   -JB 
        '                                           -   -                               
        '                                          -   -                                
        '                               `   . .  -  -                                   
        '                                .*` .* ;`*,`.,                                 
        '                                 `, ,`.*.*. *                                  
        '__________________________________*  * ` ^ *____________________________       

    End Sub

    Private Sub btnCopyToClipboard_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCopyToClipboard.Click
        Clipboard.SetText(Me.txtExceptionText.Text)
        Me.btnCopyToClipboard.Text = "Copied Error Message to Clipboard"
    End Sub
End Class
