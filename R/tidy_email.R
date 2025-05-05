#' @noRd
tidy_email <- function(
    x,
    include_body = c("text", "html", "preview", "none"),
    ...) {
  if (!inherits(x, "ms_outlook_email")) {
    stop("`x` must be an `ms_outlook_email` object", call. = FALSE)
  }

  include_body <- match.arg(include_body)

  # ---- helpers ---------------------------------------------------------
  collapse_names <- function(lst) {
    if (length(lst) == 0) {
      return(NA_character_)
    }
    paste(
      vapply(
        lst,
        function(rec) rec$emailAddress$name %||% rec$email_address$name,
        ""
      ),
      collapse = "; "
    )
  }

  collapse_addresses <- function(lst) {
    if (length(lst) == 0) {
      return(NA_character_)
    }
    paste(
      vapply(
        lst,
        function(rec) rec$emailAddress$address %||% rec$email_address$address,
        ""
      ),
      collapse = "; "
    )
  }

  collapse_categories <- function(lst) {
    if (length(lst) == 0) {
      return(NA_character_)
    }
    paste(
      lst,
      collapse = "; "
    )
  }

  safe_body_text <- purrr::safely(function(msg) {
    body_ct <- msg$properties$body$contentType # "HTML" or "Text"
    body_raw <- msg$properties$body$content
    if (is.null(body_raw) || body_raw == "") {
      return(NA_character_)
    }

    if (body_ct == "html" && include_body == "text") {
      tryCatch(
        {
          body_raw |>
            rvest::read_html() |>
            rvest::html_text2()
        },
        error = function(e) {
          # Fallback if HTML parsing fails
          gsub("<.*?>", "", body_raw) # Simple tag removal
        }
      )
    } else if (body_ct == "html" && include_body == "html") {
      body_raw # keep as-is
    } else if (body_ct == "text") {
      body_raw # plain text already
    } else {
      # fallback
      NA_character_
    }
  })

  props <- x$properties

  # ---- attachments -----------------------------------------------------
  atts <- x$list_attachments(n = Inf) # cheap metadata call
  att_names <- if (length(atts)) {
    paste(vapply(atts, \(a) a$properties$name, ""), collapse = "; ")
  } else {
    NA_character_
  }

  # Store attachment objects for later access if needed
  attachments <- if (length(atts)) {
    atts
  } else {
    list()
  }

  # ---- body ------------------------------------------------------------
  body <- switch(include_body,
    none = NA_character_,
    preview = props$bodyPreview, # 255 chars from Graph
    {
      res <- safe_body_text(x)
      res$result
    } # text or html branch
  )

  # ---- body_with_header ------------------------------------------------
  # Compose header string
  header_lines <- c(
    paste0("From: ", collapse_names(list(props$from)), " <", collapse_addresses(list(props$from)), ">"),
    paste0("To: ", collapse_names(props$toRecipients), if (!is.na(collapse_addresses(props$toRecipients))) paste0(" <", collapse_addresses(props$toRecipients), ">") else ""),
    if (length(props$ccRecipients) && !is.na(collapse_addresses(props$ccRecipients))) {
      paste0("Cc: ", collapse_names(props$ccRecipients), " <", collapse_addresses(props$ccRecipients), ">")
    } else {
      NULL
    },
    paste0("Date: ", as.character(props$sentDateTime)),
    paste0("Subject: ", props$subject)
  )
  header_str <- paste(header_lines, collapse = "\n")
  body_with_header <- if (!is.na(body)) paste(header_str, "\n\n", body) else header_str

  # ---- build tibble ----------------------------------------------------
  tibble::tibble(
    id = props$id,
    internet_message_id = props$internetMessageId %||% NA_character_,
    conversation_id = props$conversationId %||% NA_character_,
    sent_datetime = props$sentDateTime,
    received_datetime = props$receivedDateTime,
    subject = props$subject,
    from = collapse_addresses(list(props$from)),
    from_names = collapse_names(list(props$from)),
    to = collapse_addresses(props$toRecipients),
    to_names = collapse_names(props$toRecipients),
    cc = collapse_addresses(props$ccRecipients),
    cc_names = collapse_names(props$ccRecipients),
    bcc = collapse_addresses(props$bccRecipients),
    reply_to = collapse_addresses(props$replyTo),
    importance = props$importance,
    categories = collapse_categories(props$categories),
    is_read = props$isRead,
    has_attachments = props$hasAttachments,
    attachment_names = att_names,
    attachments = list(attachments), # Store actual attachment objects
    raw_email = list(x), # Store raw email object for access to original properties
    body = body,
    body_with_header = body_with_header
  )
}
