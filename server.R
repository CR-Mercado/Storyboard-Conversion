
library(shiny)
library(stringi)
library(officer)

# Define server logic required to draw a histogram
shinyServer(function(input, output,session) {
   
  #read the file using library(officer)
  observeEvent(
    eventExpr = input[["submit"]],
    handlerExpr = {
      print("Running")
    }
  )
  #this starts once the "submit" button as been hit
  output$main.content <- eventReactive(input$submit,{ 
  
    #input the uploaded file
  inFile <- input$the.file
   # read it from its location [NOTE: uploaded files create a data frame of columns: name,size, type, datapath]  in datapath 
  the.doc <- read_docx(inFile$datapath)
  content <- docx_summary(the.doc)
  
  # convert to a ppt 
  if(input$makeppt == TRUE){make.ppt(content = content)}
  
  # create individual slide .txt files 
  if(input$makeaudio == TRUE){make_audio_files(content = content)}
  
})
  session$onSessionEnded(function() {
    stopApp()
    q("no")
  })
  })
  

########################################### Functions to be used 
###########################################

#This subset of the content is for narration -  retaining only the style and text columns 
make_audio_files <- function(content){ 
  # only keeping slide name and narration 
  audio_content <- content[content$style_name %in% c("heading 1","Quote"),c(3,4)]
  
  
      # Create an index of the slides by title, i.e. index of <Heading 1> 
    title.index <- grep("heading 1",audio_content[,1]) # this reads the style column and locates "heading 1" 
    
    num.files <- length(title.index)
      
    ### We're going to put EVERYTHING between two headings into a single "content" category and put it inside the text file
    
    #initialize j      
    j <- 1
         # for each slide 
    
    dir.create("audio_files_here")
    setwd("./audio_files_here")
    while(j < num.files){
             #this says get all the numbers between two heading index 
        body.index <- title.index[j]:title.index[j+1]
        #removing the two headers:   i.e. 2:6 -> 2 3 4 5 6 --> remove 2 (the 1st #) and 6 (the last #) -> 3 4 5 
        body.index <- body.index[-c(1,length(body.index))]
  
        # combine all the body text into a single string, separated by " "
         body.txt <- paste(audio_content[body.index,2],collapse = "")
         body.txt <- gsub(".",". ",body.txt, fixed = TRUE)  # fixes the following:   The sentence.Has no space after the.period.
         body.txt <- gsub(".  ",". ",body.txt,fixed = TRUE)  # fixes: The sentence.  Has two spaces.  after the period.
         body.txt <- gsub("?","? ",body.txt, fixed = TRUE) #fixes: Does this sentence have a question mark error?Yes it does.
         body.txt <- gsub("?  ","? ",body.txt, fixed = TRUE) # fixes:  Does this sentence have an extra space?  Yes it does.
         
    # split the captions into <=330 char chunks 
        body.txt <- caption.splitter(body.txt)
         
    # write the chunks into a .txt named after the slide name and removing punctuation from the slide name (invalid name char for txt)
        curr.slide.name <- gsub("[[:punct:]]"," ",audio_content[title.index[j],2]) 
       write(body.txt, file = paste0(j,"-",curr.slide.name,".txt"))
         j <- j+1                
    }
     
        
    setwd("..")
    
}



caption.splitter <- function(the.text, max.char = 330, separator = "          "){ 
  # This function takes text and adds the separator (default is 10 spaces "          ") 
  # in order to visually split the caption text into chunks (maximum 330 characters per chunk default)
  # In our e-learning environment the caption blocks have a max number of characters that fit without expanding the box. 
  # We would like this to be clearly visible from the text file - while not interrupting the auto-narration - for at minimum easy copy/paste
  # and at best location of breaks and importation into the caption translation file box
  
  
  
  #get the location of all the spaces in the text box  
  original.space.positions <- gregexpr(" ",the.text) #this includes some metadata
  original.space.positions <- original.space.positions[[1]][1:length(original.space.positions[[1]])]  
        #this gets an integer vector of the space positions
  
 
  
  #set an empty vector for all of the chunks
  chunk <- NULL
  
  #set counters
  i <- 0
  j <- 1
  
  
  
  #while i is less than the highest space position
  while(i < max(original.space.positions)){   
    print("loop counting")
    #the j'th chunk is the substring from the character after i'th space (i.e. the word) to the farthest space allowed by max.char  
    chunk[j] <- substring(the.text,i+1, 
                          max(original.space.positions[original.space.positions <= (i + max.char)]))
    
    #then change i to that farthest space allowed, essentially we are rolling through the text, separating chunks of text by as much as possible without 
    #cutting up a word. 
    i <- max(original.space.positions[original.space.positions <= i + max.char])
    
    #move to the next chunk counter
    j <- j+1 
    
  }
  
  #after the while loop there is one chunk left. The final i all the way to the last character which can be stored in the j'th chunk.
  #NOTE: if somehow the last singlewordinthedocumentwasover max.char in length then this process would break. So keep max.char longer than the longest word in 
  # the document
  
  chunk[j] <- substring(the.text, i, nchar(the.text)) # i is at the last legal space so this goes from there to the end. 
  
  
  # The previous process ALWAYS makes the last word into its own chunk. 
  # This tests the j-1'th chunk to see if it is short enough to fit that last word
  # if there's room for it, it will combine the two and remove the gap.  
  
  if( (max.char - nchar(chunk[j-1])) 
      > nchar(chunk[j])
  ){
    chunk[j-1] <- sub("   "," ",paste(chunk[j-1],chunk[j], collapse = "")) #swap out 3 spaces for 1, because each chunk holds extra spaces 
    # but we are overriding that. 
    chunk <- chunk[1:(j-1)]
  }
  print(paste0("number of chunks: ",length(chunk))) #print the number of chunks 
  
  # Return the spaced out text
  return(paste(chunk, collapse = separator))
  
}
##############################

chunks.counter <- function(chunked.output,name.of.file,separator = "          "){
  # this function takes a chunked output and counts how many chunks there are
  # it spits out the name of the file and the number of chunks required 
  # 
  #
  numchunks <- length(strsplit(chunked.output, paste0(separator," "))[[1]])
  
  return(c(name.of.file, numchunks)) #returns a class character vector. Which is easy to rbind 
}

###############################

splits.to.table <- function(chunked.output, separator = "          "){ #reuse the separator from previous function - default 10 
  # 
  # it then reverses the order of the chunks and puts it into a matrix of 1 column
  #
  numchunks <- length(strsplit(chunked.output, paste0(separator," "))[[1]])  # this is copied from the previous function but it is not used
  # as a return here 
  #it does output as integer
  chunk.matrix <- matrix(strsplit(chunked.output,paste0(separator, " "))[[1]][numchunks:1]) 
  
  # Be careful changing the default separations because this sub will have separator # of spaces / 2 on all non chunk 1. 
  # this shrinks extra spaces down to 1 by repeatedly removing back to back spaces
  chunk.matrix <- gsub("  ","",chunk.matrix)
  
  return(chunk.matrix)
}


############################ 

make.ppt <- function(content){ 
  # This function takes a subset of a word document (style name is <h1> OR <h2> or <h6> and creates a ppt from it)
  # everything between two <h1> goes into the <h1> slide before it 
  # various defaults were taken from the template names 
  # use layout_properties(read_pptx("TEMPLATE FILE NAME HERE")) #to see the names, type, etc info including offsets.
  #
  
  #empty powerpoint from the template - included in shiny file
  the.pres <- read_pptx("Lectora_SlideMasterTemplate_v6.pptx") 
  
  # This subset of the content is for screen text through ppt - retaining only the style and text columns 
  ppt_import_content <- content[content$style_name %in% c("heading 1", "heading 2", "heading 6"),c(3,4)]
  
  # Create an index of the slides by title, i.e. index of <Heading 1> 
  slide.title.index <- grep("heading 1",ppt_import_content[,1]) # this reads the style column and locates "heading 1" 
  
  ### We're going to put EVERYTHING between two headings into a single "content" category and put it inside the slide 
  
  num.slides <- length(slide.title.index)

  #initialize j 
  j <- 1
 
   # While the number of slides has not been reached: 
  while(j < num.slides){
    
    #add a blank slide  -layout = Lesson slide IF the title BEGINS WITH "Lesson" else - title + content
    the.layout <- ifelse(grepl("^Lesson",ppt_import_content[slide.title.index[j],2]),"Lesson_Slide", "1_Title + Full Content")
    the.pres <- add_slide(the.pres, layout= the.layout, master = "Office Theme") 
    
    # make the title equal to the heading one that we are currently on 
    the.pres <- ph_with_text(the.pres, type = "title", str = ppt_import_content[slide.title.index[j],2])
    
    #this says get all the numbers between two heading index 
          body.index <- slide.title.index[j]:slide.title.index[j+1]
    #removing the two headers:   i.e. 2:6 -> 2 3 4 5 6 --> remove 2 (the 1st #) and 6 (the last #) -> 3 4 5 
          body.index <- body.index[-c(1,length(body.index))]
    
    # combine all the body text into a single string, separated by " "
     body.txt <- paste(ppt_import_content[body.index,2],collapse = " ")
   the.pres <-ph_with_text(the.pres, type = "body", str = body.txt)
    
  # count up to the number of slides  
  j <- j + 1
          }
print(the.pres, target = paste0(j,"numslides-pleaseRename",".pptx"))
invisible()
}

#####################



  
  