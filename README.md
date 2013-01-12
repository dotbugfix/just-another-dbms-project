#just-another-dbms-project

Back in the day, we had to make a (mini)-project of sorts to showcase our lack of skills in a dodgy subject called DBMS. Our tools were neanderthal and would put a vintage stop-motion animator to his knees: Microsoft Visual Basic 6.0 and a cherry-picked edition of Oracle. So if you're an under-grad student, on the verge of a term-end, looking for a quick jack-up on your overdue submissions, you're in the right place.

##Trailer
For guys who would rather sit down, lay back & watch a movie than pick up that controller and game on, here's your gift:

###[Entertainment Resort Management System - In a Nutshell](https://github.com/dotbugfix/just-another-dbms-project/blob/master/Documentation/EntResort%20Documentation.pdf)

P.S.: Shortly after browsing through the 19-page PDF, you'll realize that the *nutshell* was actually a metaphor for *endless saga of perpetual demonstration and whimsical quotes* and the only thing that it has in common with the notion of a *trailer* is it's mildly entertaining.

##Purpose
Given the very specific nature of these projects, it's easy to wander off in the pool of resources that the Internet throws at you.  Even the maestros of painting often prefer a subject to look at. Here is your subject: take it as a 'model solution', if you will; keeping in mind the obscure properties of model solutions in general: they will seldom make you top the class, but they a guaranteed _never_ to raise an eyebrow, much like an underdog (with a degree).

##Package
I have no idea why you're still reading this, but maybe it because you've not spotted the shiny 'Download' button anywhere. Anyway, here's the deal:

1. [EntResort (Portable Setup).zip](https://github.com/dotbugfix/just-another-dbms-project/blob/master/EntResort%20%28Portable%20Setup%29.zip)
    * This is a standalone next-next-finish style setup that you can take out for a test ride
    * You'll need a jumpstart for the credentials and stuff, so look [here](https://github.com/dotbugfix/just-another-dbms-project/blob/master/Documentation/EntResort%20Documentation.pdf) for some clues

2. [VB Project (MS Access)](https://github.com/dotbugfix/just-another-dbms-project/tree/master/VB%20Project%20%28MS%20Access%29)
    * This is an MS Access edition of the project. It's simply a port for people who have just woken up and are yet to install Oracle on their system. The connection strings have been morphed to use a [`db.mdb`](https://github.com/dotbugfix/just-another-dbms-project/blob/master/VB%20Project%20%28MS%20Access%29/Database/db.mdb) file in the current directory.

3. [VB Project (Oracle)](https://github.com/dotbugfix/just-another-dbms-project/tree/master/VB%20Project%20%28Oracle%29)
    * This is the original project, designed to work with a local installation of Oracle 9i (though anything higher should work, but is probably not prescribed in the course)
    * Database dumps can be found under [`/Database/EXPDAT.EXP`](https://github.com/dotbugfix/just-another-dbms-project/blob/master/VB%20Project%20%28Oracle%29/Database/EXPDAT.DMP), which you can import after a couple of futile attempts at first finding and then getting the import utility wrapped around your head.

4. [Documentation](https://github.com/dotbugfix/just-another-dbms-project/blob/master/Documentation/EntResort%20Documentation.pdf)
    * I've tossed in a complimentary ['sample project report'](https://github.com/dotbugfix/just-another-dbms-project/blob/master/Documentation/Project%20Report%20%28Entertainment%20Resort%29.doc) with the 'model solution', mostly in a vague attempt to colour inside the lines, if you know what I mean.
    * If you don't know what I mean, maybe the [artiste's imitation](https://github.com/dotbugfix/just-another-dbms-project/blob/master/Documentation/EntResort%20Documentation.pdf) would appeal to you: it's a crash-course in making your first college (mini)-project that wades through my team & mine meandering experiences.

##Quickstart Guide
So you've jumped right to this part; I can understand your anxiety in the wee hours of submission day. I'll get straight to the point, here's what you need to get this thing up and running before your noodles boil out:

1. Get the [`EntResort (Portable Setup).zip`](https://github.com/dotbugfix/just-another-dbms-project/blob/master/EntResort%20%28Portable%20Setup%29.zip) file and run Setup.exe from it.
2. You'll find an entry for "Entertainment Resort" in your Start menu (or whatever it is called these days)

    i. First open the [`EntResort Documentation.pdf`](https://github.com/dotbugfix/just-another-dbms-project/blob/master/Documentation/EntResort%20Documentation.pdf) file and turn to *page 6: "Factory Settings"*, where you'll find all the preset data and credentials for going past the welcome screen.

    ii. If you really need a helping hand, turn to *page 7: "Guided Tour"*.


Now, after you are satisfied by the potency of the frontend, you probably want to dive into the code. That's rather easy (to look at):

1. Get either of the [`VB Project (MS Access)`](https://github.com/dotbugfix/just-another-dbms-project/tree/master/VB%20Project%20%28MS%20Access%29) or [`VB Project (Oracle)`](https://github.com/dotbugfix/just-another-dbms-project/tree/master/VB%20Project%20%28Oracle%29) directories and fire up the `/Source Code/EntertainmentResort.vbp` file in your collector's edition of Visual Studio 6.0
2. Anything else that you may need is elaborated on *page 14: "Source Code"* of the [documentation](https://github.com/dotbugfix/just-another-dbms-project/blob/master/Documentation/EntResort%20Documentation.pdf).

##For the older and wiser souls (a.k.a. Contributing)
Yes, this is pretty naive stuff. And yes, I know you're winching at the folder hierarchy and thinking that I should've probably made separate branches or something. And that binary, cheekily named [`EntResort (Portable Setup).zip`](https://github.com/dotbugfix/just-another-dbms-project/blob/master/EntResort%20%28Portable%20Setup%29.zip) is probably bothering you to no end. But if you sit down, take another sip of whatever it is you're drinking, and ponder over the wiggly movements of our meek target audience, I'm sure you'll understand. And trust me, if I even dared to *think* about approaching a cloud-based continuous integration site with this source code [sic: VB6], I'll be banished forever (just the Google search would maybe block my account for a week or something). So if it's a late hour and you've just completed the last instalment of that RPG and are basking in glory, you might even want help calming down that Adrenelene rush: have a look at [some issues](https://github.com/dotbugfix/just-another-dbms-project/issues) to kill time peacefully.

Meanwhile, check out my [other repos](https://github.com/dotbugfix?tab=repositories)!

##Licensing
This will probably be taken care of after I license some of my other projects first (and also probably by the guy who takes care of that).

##Credits
###My Project Team
Apart from the few weeks of sleepless nights, this project is the outcome of these amazing people who made my days sleepless too; who taught me that even if you had all the time in the world and decided to do every damn thing by yourself, it wouldn't be as good if you had a team to do it; and who seem to be leading much healthier lives since the only link I have here is to their Facebook profiles:
* [Amruta Gandhi](https://www.facebook.com/amruta.gandhi.3)
* [Amit Dikkar](http://www.facebook.com/amit.dikkar)

###Invaluable resources in disguise
This `README.md` file was made with great help from [GitHub Markup Preview](http://dfilimonov.com/github-markup-preview/), whose repo you can find [here](https://github.com/petethepig/github-markup-preview). I sincerely hope it will encourage scores of developers to leverage the [Github flavored markup](https://github.com/github/markup/) for their own documentations.

##A conservative timeline
This project was made sometime back in 2011. I do not know if it will stand the test of time; and I would be the gladdest person it it doesn't: for this entails that the prescribed syllabus to students of age has been finally revised to suit the day (in which case this would probably count as an *'History of Information Technology'* article for whatever class it is called these days!
It still holds ground as one of the first collaborative projects that I've done, but I quickly realize that GitHub might be skeptical of that!