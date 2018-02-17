library(shiny)

shinyUI(fluidPage(
  titlePanel("EWARS"),
  
  sidebarLayout(
    sidebarPanel(
      
      # https://countrycode.org/
      selectInput('country_code', 'country code', 
                  c("XX"),
                  selected="XX", width="35%"),
      textInput("psswd", "8 digit password", "password"),
      checkboxInput("generating_surveillance_workbook", "Upload workbooks to the database", value=FALSE),
      tags$hr(),
      
      fileInput('dat', 'choose file to upload', accept=c('.xls','.xlsx')),
      textInput("original_data_sheet_name", "Choice of sheet name for the original data", "Sheet1"),
      tags$hr(),
      
      helpText("THE USER CAN BEGIN THE SETTINGS/ CALIBRATIONS FROM THIS POINT"),
      
      #br(),
      textInput("stop_runin", "Specify the year/week where the run-in period stops", "201252"),
      #br(),
      
      #br(),
      textInput("run_per_district","Enter the corresponding code(s) for the District/ Municipality you wish to analyse.If you chose more than one, use comma between each District/ Municipality", "3,15"),
      #br(),
      textInput("population", "Enter the variable name which represents the annual total Population of the corresponding district/ municipality", "population"),
      # br(),
      textInput("number_of_cases", "Enter the variable name which represents the weekly number of outbreak", "weekly_hospitalised_cases"),
      # br(),
      numericInput("outbreak_week_length","Choice of the number of outbreak weeks to declare outbreak period/stop outbreak", value = 3),
      #br(),
      textInput("alarm_indicators", "Enter the alarm indicator(s) you wish to analyse. If you chose more than one, use comma between each alarm indicator", "meantemperature,rainsum"),
      #br(),
      numericInput("alarm_window","Enter the window size (e.g 2) for measuring the mean alarm indicator including current week", value = 2),
      # br(),
      numericInput("alarm_threshold","Enter the choice of alarm threshold: is the value from which the alarm signal can be declare (e.g 0.12)", value = .12),
      #br(),
      
      checkboxInput("graph_per_district", "Specify graph per district/municipality option",value =TRUE),
      #br(),
      numericInput("season_length","Season length", value = 52),
      # br(),
      numericInput("z_outbreak","Enter the multiplier of the standard deviation to vary the endemic channel within the evaluation period (e.g. 1.25)", value = 1.25),
      # br(),
      numericInput("outbreak_window","Enter the choice of outbreak window (e.g 3)", value = 3),
      #br(),
      numericInput("prediction_distance","Enter the desired choice of distance between current week and target week to predict an outbreak signal (e.g 2)", value = 2),
      # br(),
      numericInput("outbreak_threshold","Enter a choice of cut-off value to define the outbreak signal (e.g 0.25)", value = 0.25),
      # br(),
      
      checkboxInput("spline", "Spline option",value =FALSE),
      br(),
      actionButton("goButton", "Run!")
      
      
    ),
    mainPanel(
      tabsetPanel(
        tabPanel("Runnin Period", uiOutput("plot1")),
        tabPanel("Evaluation Period", uiOutput("plot2")),
        tabPanel("Runnin Evaluation Period", uiOutput("plot3")),
        tabPanel("Sensitivity/Specificity", tableOutput("table1")),
        #tabPanel("Workbooks", uiOutput("workbooks")),
        tabPanel("Help", 
                 tags$br(),tags$br(),
                 tags$div(
                   HTML(paste(tags$strong("Early Warning and Response System for Dengue Outbreaks: Operational Guide using the web-based Dashboard"), 
                              tags$a(href="http://globalminers8973.cloudapp.net/EWARS-R_dashboard.pdf", target="_blank", tags$br()," [download the pdf file here]"),
                              sep = ""))
                 ),
                 tags$br(),tags$br())
      )
    ) 
    
  )
))

