% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/data.functions.R
\name{rerefout}
\alias{rerefout}
\title{Change reference treatment for results vs a reference}
\usage{
rerefout(
  ref = "Pill placebo",
  trtcodes,
  es = "or",
  level = "treat",
  bugsres,
  model.scale = "natural"
)
}
\arguments{
\item{ref}{A character object for the reference \strong{treatment} (not the reference class)}

\item{trtcodes}{A data frame of treatment/class names and codes}

\item{es}{A character object corresponding to the name of the node/parameter in WinBUGS
for the treatment/class effects vs every other treatment/class}

\item{level}{Whether \code{rerefout()} should be at the treatment (\code{"treat"}) or class (\code{"class"})
level}

\item{bugsres}{A data frame of BUGS posterior summaries}

\item{model.scale}{Indicates whether treatment effects are on \code{"log"} (e.g. OR, RoM) or
\code{"natural"} (e.g. mean difference) scale. If left as \code{NULL} (the default) this will be
inferred from the data.}
}
\description{
Change reference treatment for results vs a reference
}
