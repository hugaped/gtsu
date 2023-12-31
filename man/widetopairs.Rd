% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/data.functions.R
\name{widetopairs}
\alias{widetopairs}
\title{Converts a wide data frame into contrasts}
\usage{
widetopairs(wide.df, output = "studies")
}
\arguments{
\item{wide.df}{A BUGS-ready data frame in wide format (one row per study) containing \code{study} and \code{narms}
variables.}

\item{output}{Can take either \code{"studies"} to indicate all study-level pairs should be generated, or
\code{"contrasts"} to indicate that only treatment/class level contrasts should be generated.}
}
\description{
Converts a wide data frame into contrasts
}
