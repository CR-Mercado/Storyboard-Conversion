# This app reads a word document
# creates a ppt from the <H1> <H2> and <H6> content
# creates separate text files for narration .mp3
# creates a table dictating the # of caption blocks needed for each slide 
# creates a character vector of all the caption chunks to replace the Lectora html file later [in reverse per slide order]
# creates a word document of <h1> and narration chunked with large spaces. 
#


library(shiny)


shinyUI(fluidPage(
  
  # Application title
  titlePanel("Storyboard Derivatives"),
  
  # Sidebar   
  sidebarLayout(
    sidebarPanel(
      ## Select a file --- 
      h6("This Program turns a formatted word doc into a ppt (for lectora import) and separate narration .txt files for each slide. 
         The txt files are chunked by large spaces to easily identify legal caption sizes (330 chars in lectora)."),
      h6("The program expects an empty slide template as the last page. Please ensure the ppt is not missing a slide and all audio files are there."),
      h6("Note: The program will show you where the files are upon completion: under the names audio_files_here and N_numslides_PleaseRename where N is 
the number of slides in the powerpoint. Please move & rename both IMMEDIATELY to the correct location and name. The .txt will need to be run through 
the GOOGLE Text to Speech API - a separate app is available with our custom JSON authentification for that purpose."),
      h6("Always match up the number of slides and audio files to what you expected. If you get any weird errors,
         it is possible that a slide has No <Quote> styled narration. Please insert a blank space with style <quote> in the storyboard in that location
         -(you can tell which one because the app will stop making txt files at the slide before the error slide)"),
      h6("Always move files immediately upon finishing the program, before running on a new storyboard"),
      fileInput("the.file","Choose Word Docx",
                multiple = FALSE,
                accept = ".docx",
                buttonLabel = "Browse"),
      
      actionButton(inputId = "submit",label = "Run Program"),
      checkboxInput("makeppt","Make a PPT", value = TRUE),
      checkboxInput("makeaudio","Make audio txt files", value = TRUE)
    ),
    
    # Main Panel
    mainPanel(
       textOutput("main.content")
    )
  )
))
