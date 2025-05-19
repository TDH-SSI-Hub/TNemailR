

#' Checks and walkthrough to enable RDCOMClient emails.
#'
#' @param force Ignore previous RDCOMClient installation
#' @param repo Github repository to use. omegahat or bschamberger.
#'
#' @return T/F
#' @export

email_setup<-function(force=F, repo='bschamberger'){
  unloadNamespace("RDCOMClient")
  if(force) utils::remove.packages("RDCOMClient")

  if('RDCOMClient' %in% rownames(utils::installed.packages())){
    message('RDCOMClient already installed.\nIf you have problems sending emails, uninstall RDCOMClient and try again.')
  return(T)
    }

if('devtools' %in% rownames(utils::installed.packages())){
  message('devtools package already installed')
}else{
  message('installing devtools package')
  utils::install.packages('devtools')
}



  #if(pkgbuild::check_rtools()){
  #  message('rtools already installed')
  #}else{
  #  rt_ver<-substr(paste0(R.Version()$major,R.Version()$minor),1,2)
  #  message(paste0('Need to install rtools ',rt_ver,'. Follow prompts in dialog box.'))
  #  pkgbuild::check_build_tools()
  #  return(F)
  #}

  valid_repo<-c('omegahat','bschamberger')

  if(tolower(repo) %in% valid_repo){

  #message(paste0('Installing RDCOMClient from ',repo,'/RDCOMClient'))
  tryCatch(devtools::install_github(paste0(repo,'/RDCOMClient'))
           , warning=function(w){
             message('RDCOMClient installation failed')
             return(F)
             }
           , error=function(w){
             message('RDCOMClient installation failed')
             return(F)
             }
           , finally = function(w) {
             message('RDCOMClient installed')
             return(T)
           }
  )

  }else{
    message(paste0('repo parameter not valid. Use one of the following: ',paste0(valid_repo, collapse=', ')))
  }

}


#' Draft and possibly send an Outlook email.
#'
#'
#' @param to Main recipients
#' @param subject Email subject
#' @param body Email body. Searches for html line break to decide if body is html
#' @param from Alternate email to send on behalf of
#' @param attach Character vector of files to attach
#' @param cc Email cc
#' @param bcc Email bcc
#' @param visible Open the email in the viewer
#' @param check_ooo Check for and remove recipients who are out of office the entire day. If there are no more main recipients, the email is not sent. Requires open Outlook application.
#' @param send Send the email
#' @param signature Attach signature to end of email. Requires open Outlook application.
#'
#' @return T/F
#' @export
email_draft <- function(to='', subject='', body='', from = NA, attach = c(), cc = c(), bcc = c(), visible = T, check_ooo = F, send = F, signature=F) {
  if(require('RDCOMClient')){

    check_error<-tryCatch(getCOMInstance('Outlook.Application', force=F), error=function(e){
      return(e)
    })

    if(!"COMIDispatch" %in% unlist(class(check_error))){
      message('No Outlook instance open; signature and OOO check disabled')
      check_ooo <- F
      signature <- F
    }
    if(!exists('outApp')){
      outApp<<-RDCOMClient::COMCreate("Outlook.Application", existing = F)
    }

    #if (visible) {
      outMail <<- outApp$CreateItem(0)
    #}else{
    #  outMail <<- outApp$CreateItem(0)
    #}

    # Send the message from an alternate account
    if (!is.na(from)) {
      outMail[["sentonbehalfofname"]] <- from
    }

    if (visible) {outMail$Display()}

    if(signature){
      inspector<-outMail$GetInspector()
      signaturetext <- outMail[["HTMLBody"]]
      #inspector$Close(1)
    }

    for (r in c("to", "cc", "bcc")) {
      if (length(get(r)) > 0) {
        outMail[[r]] <- paste0(get(r), collapse = ";")
      }
    }

    if (check_ooo) {
      remove_list <- c()
      for (i in seq_len(outMail[["recipients"]]$Count())) {
        tryCatch({
          time_string <- substr(outMail[["recipients"]]$item(i)[["AddressEntry"]]$GetFreeBusy(as.character(Sys.Date()), 60, T) , 1, 24)
          if (time_string == paste0(rep(3, 24), collapse = "")) {
            message(paste0(outMail[["recipients"]]$item(i)[["Name"]], " is OOO"))
            remove_list <- c(remove_list, i)
          } else {
            #message(outMail[["recipients"]]$item(i)[["Name"]])
          }
        }, error=function(e) message(paste0('Error checking OOO status for ',outMail[["recipients"]]$item(i)[["Name"]])))
      }

      for (i in remove_list) {
        outMail[["recipients"]]$Remove(i)
      }
    }



    outMail[["subject"]] <- subject


    if (length(attach) > 0) {
      for (i in attach) {
        if (!grepl(getwd(), i)) i <- paste0(getwd(), "/", i)
        tryCatch(
          {
            outMail[["attachments"]]$Add(i)
          },
          error = function(e) message(paste0("Failed to attach ", i))
        )
      }
    }

    if (grepl("<br>|<tr>",body)) {
      outMail[["HTMLBody"]] <- body
    } else {
      outMail[["body"]] <- body
    }


    if(signature ){
      if(!is.null(signaturetext)){
      outMail[["HTMLBody"]] <- paste0(outMail[["HTMLBody"]],signaturetext)
      }else{
        message("Error generating signature")
      }
    }





    if (send) {

      #if (outMail[["to"]] == "") {
      #  message("Email has no main recipient; email discarded")
      #  outMail$Close(1)
      #  return(F)
      #}


      return(outMail$Send())
    }else{
      return(NA)
    }
  }else{
    message('You need to install RDCOMClient to send emails')
    return(F)
  }
}


#' Send an Outlook email. Wrapper around email_draft()
#'
#'
#' @param to Main recipients
#' @param subject Email subject
#' @param body Email body. Searches for html line break to decide if body is html
#' @param from Alternate email to send on behalf of
#' @param attach Character vector of files to attach
#' @param cc Email cc
#' @param bcc Email bcc
#' @param visible Open the email in the viewer
#' @param check_ooo Check for and remove recipients who are out of office the entire day.
#'  If there are no more main recipients, the email is not sent. Requires open Outlook application.
#' @param send Send the email
#' @param signature Attach signature to end of email. Requires open Outlook application.
#'
#' @return T/F
#' @export
email_send <- function(to, subject, body, from = NA, attach = c(), cc = c(), bcc = c(), visible = F, check_ooo = F, send = T, signature=F) {
  email_draft(to=to,
              subject=subject,
              body=body,
              from=from,
              attach=attach,
              cc=cc,
              bcc=bcc,
              visible = visible,
              check_ooo=check_ooo,
              send = send,
              signature=signature
              )
}




