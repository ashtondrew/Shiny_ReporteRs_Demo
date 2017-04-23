library(ggplot2)
library(magrittr)
library(dplyr)
library(ReporteRs)

invites <- tibble(Leaders = c(rep("LeaderX", 8), rep("LeaderY", 6)),
                      Subjects = paste(rep("Subject", 14), 1:14))
responses <- tibble(Leaders = c(rep("LeaderX", 8), rep("LeaderY", 6)),
                        Subjects = paste(rep("Subject", 14), 1:14),
                        Var1 = rep(c("Yes","No", "Yes"), length=14), #rbinom(14, 1, 0.75),
                        Var2 = runif(14, 0, 1),
                        Var3 = runif(14, 0, 1))
dat <- merge(invites, responses, by = c("Leaders", "Subjects"))

piedat <- dat %>%
  group_by(Var1) %>%
  summarize(CntVar1 = n()) %>%
  ungroup()

ui <- fluidPage(
  # The app contains fileInput, uiOutput and dataTableOutput code to upload files, select subject for reporting, and view summary tables.  For this reproducible code, I provided two dataframes in the server code and substitute in a simple selectInput and tableOutput.
  tableOutput("StatusSummary"),
  selectInput("LeaderID", "Select a leader to generate a report:", unique(dat$Leaders)),
  downloadButton("report", "Save Report")
)

server <- function(input, output) {
  output$StatusSummary <- renderTable({
    dat
  })
  output$report <- downloadHandler(
    # the filename to use
    filename = function() {
      paste0("Demo_", input$LeaderID , "_",Sys.Date(),".pptx")
    }, 
    # the document to produce
    content = function(file){
      # use custom template
      doc <- pptx(template = 'Project_Template.pptx')
      # Add a slide by first selecting a layout from your Master Slide layout styles
      doc <- addSlide( doc, slide.layout = 'Project_TitleSlide' )
      # Then fill in the relevant elements for that slide
      doc <- addTitle( doc, input$LeaderID)
      doc <- addSubtitle( doc, paste(Sys.Date()))
      
      # Add slide with figures - notice difference between adding base plots and ggplots
      doc <- addSlide( doc, slide.layout = 'Project_Scatter&Pie' )
      doc <- addTitle( doc, "Here is some title text")
      PlotA <- ggplot(dat, aes(Var2,Var3))+geom_point()
      doc <- addPlot( doc, function() print(PlotA))
      doc <- addPlot( doc, function() pie(piedat$CntVar1, labels=piedat$Var1))
      doc <- addParagraph (doc, "This is a comment.")
      doc <- addDate( doc )
      doc <- addPageNumber( doc )
      
      writeDoc( doc, file )
    } # end of report  content function
  ) # end of downloadHandler function
  
} # end of server function

# Run the application 
shinyApp(ui = ui, server = server)