# xlsx-tools
# assist operations w/ Excel
# by Dr John Van Camp  5/01/2018

require(jvctools); defl()
require(openxlsx)
require(glue)

xlstyle<-function(wb,ws,rows,cols,style=NULL) {
  #apply string of styles to worksheet
  #align al ar ac at am ab (left,right,center,top,middle,bottom)
  #text t.bold t.small t.norm t.big t.red t.blue t.black 
  #num d0-d4 c0-c4 p0-p4 
  #border ba bt bb bl br (all,top,bottom,left,right)
  #fill 
  work<-unlist(strsplit(style," "))
  #browser()
  for(i in work){
    sty<-switch(i,
          "ac"=createStyle(halign = "center"),  # align left,right,center
          "al"=createStyle(halign = "left"),    #       top,middle,bottom
          "ar"=createStyle(halign = "right"),
          "at"=createStyle(valign = "top"),
          "am"=createStyle(valign = "center"),
          "ab"=createStyle(halign = "bottom"),
          "d0"=createStyle(numFmt = "##0"),     #numFmt decimal,comma,percent
          "d1"=createStyle(numFmt = "##0.0"),
          "d2"=createStyle(numFmt = "##0.00"),
          "d3"=createStyle(numFmt = "##0.000"),
          "d4"=createStyle(numFmt = "##0.0000"),
          "c0"=createStyle(numFmt = "#,##0"),
          "c1"=createStyle(numFmt = "#,##0.0"),
          "c2"=createStyle(numFmt = "#,##0.00"),
          "c3"=createStyle(numFmt = "#,##0.000"),
          "c4"=createStyle(numFmt = "#,##0.0000"),
          "p0"=createStyle(numFmt = "0 %"),
          "p1"=createStyle(numFmt = "0.0 %"),
          "p2"=createStyle(numFmt = "0.00 %"),
          "p3"=createStyle(numFmt = "0.000 %"),
          "p4"=createStyle(numFmt = "0.0000 %"),
          "ba"=createStyle(border="TopBottomLeftRight",borderStyle="medium"),
          "bt"=createStyle(border="Top",borderStyle="medium"),
          "bb"=createStyle(border="Bottom",borderStyle="medium"),
          "bl"=createStyle(border="Left",borderStyle="medium"),
          "br"=createStyle(border="Right",borderStyle="medium"),
          "t.small"=createStyle(fontSize =8),      
          "t.norm"=createStyle(fontSize =10),
          "t.big"=createStyle(fontSize =14),
          "t.blue"=createStyle(fontColour ="blue"),
          "t.red"=createStyle(fontColour ="red"),
          "t.black"=createStyle(fontColour ="black"),
          "t.bold"=createStyle(textDecoration ="bold"),
          "t.ul"=createStyle(textDecoration ="underline"),
          "t.ul2"=createStyle(textDecoration ="underline2"),
          default=NULL
    )
    if(!is.null(sty)){
      addStyle(wb, ws, style=sty,rows,cols, gridExpand = TRUE,stack = TRUE)
    }
  }
}

setcolw<-function(wb,ws,cols,widths=NULL){
  # adjust column widths in wb,ws 
  # widths a list of widths (recycle last element as needed)
  if(length(widths)>(length(cols))) { #trim if too long
    widths<-widths[1:length(cols)]
    outdata<-data.frame(col=cols,colw=widths)
  } else { if(length(cols)>length(widths)) { #fill with last element
    addcols<-length(cols)-length(widths)
    widths<-append(widths,rep(widths[length(widths)],addcols))
    outdata<-data.frame(col=cols,colw=widths) 
  } else {
    outdata<-data.frame(col=cols,colw=widths)
  }
    for(i in 1:nrow(outdata)) {
      setColWidths(wb, ws, cols=outdata[i,1], widths=outdata[i,2]) }
  } # end adjust col widths
}  


new_wb<-function(){
  wb <- createWorkbook()
  options("openxlsx.borderColour" = "#4F80BD")
  options("openxlsx.borderStyle" = "thin")
  options("openxlsx.dateFormat" = "mm/dd/yyyy")
  options("openxlsx.datetimeFormat" = "yyyy-mm-dd hh:mm:ss")
  options("openxlsx.numFmt" = "#,###,##0") ## For default style rounding of numeric columns
  modifyBaseFont(wb, fontSize = 10, fontName = "Arial Narrow")
  return(wb)
}

save_wb<-function(wb,filename){
  saveWorkbook(wb, filename,overwrite = ifelse(file.exists(filename),T,F))
}

tablim<-function(row,col,df){
  # returns limits of xl table (s.row,s.col,e.row,e.col)
  endrow<-row+nrow(df)
  endcol<-col+ncol(df)-1
  return(c(row,col,endrow,endcol))
}

testxlsx<-function(){
  df<-data.frame("letters"=LETTERS[1:9],"Nums"=c(1:9),"Rand"=1000*runif(9,1,20))
  wb<-createWorkbook()
  addWorksheet(wb,"test",gridLines = F)
  writeDataTable(wb,1,df,startRow = 2,startCol = 2,withFilter = F)
  l<-tablim(2,2,df)
  xlstyle(wb,ws=1,rows=c(3:11),cols=4,style="d1")
  save_wb(wb,"test.xlsx")
}

