# Extract Excel from Outlook into R
This is an R script that will extract excel files from Outlook and will only work on Window PC. No support available for Mac or Linux.

## Session Info
- R version 4.0.4
- readxl_1.3.1
- RDCOMClient_0.94-0

## Where to start?
1. Install R if you have not done so. Install R before Rstudio. Choose the latest version
    - [R](https://cran.r-project.org/)
2. Install Rstudio if you have not done so.
    - [Rstudio Desktop](https://www.rstudio.com/products/rstudio/download/#download)
3. Open up Rstudio and install RDCOMClient on the console tab with the following command:
    - ```install.packages("RDCOMClient", repos = "http://www.omegahat.net/R") ```
4. Next install readxl on the console tab with the following command:
    - ```install.packages("readxl") ```
5. Clone this repository or download the Zip file. The link below will guide you on how to do so:
    - [How to clone a repo on github](https://docs.github.com/en/github/creating-cloning-and-archiving-repositories/cloning-a-repository)
6. Open up my R script (main.R) and on line 12, you should see comment telling you to change the subject title. Change the subject title *Data for today* to your actual subject title.
7. Run the script by highlighting the whole script (Ctr-a) and click on "Run App" or Ctr+Enter.
8. You should be able to see the attachment being assigned to a variable call Data.
9. You can now export that data to SQL or do some data cleaning on it.

## References
- [RDCOMClient repo](https://github.com/omegahat/RDCOMClient)
