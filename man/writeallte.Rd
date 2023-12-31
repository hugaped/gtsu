% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/output.functions.R
\name{writeallte}
\alias{writeallte}
\title{Write relative effects for all comparisons in the network}
\usage{
writeallte(
  bugsdat,
  bugsres,
  umeres = NULL,
  decimals = 2,
  level = "treat",
  scalesd = NULL
)
}
\arguments{
\item{bugsdat}{An object of class \code{"bugsdat"} containing the info used to run the BUGS
model (i.e. data, treatment codes, etc.)}

\item{decimals}{The number of decimal places to which to report numerical results}

\item{level}{takes either \code{"treat"} or \code{"class"}}

\item{scalesd}{A number indicating the SD to use for back-transforming SMD to a
particular measurement scale. If left as \code{NULL} then no transformation will be
performed (i.e. leave as \code{NULL} if not modelling SMDs)}
}
\description{
Write relative effects for all comparisons in the network
}
