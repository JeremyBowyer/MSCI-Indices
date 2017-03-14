
#############################
###### MSCI FUNCTIONS #######
#############################

# These scripts will download data from MSCI's end of day history app, found here:
# https://www.msci.com/end-of-day-history?chart=country&priceLevel=41&scope=R&style=B&currency=15&size=36&indexId=119152

# Download data
msci.download <- function(countries,
			  startDate = "1969-12-29",
			  endDate = Sys.Date(),
			  priceLevel = "41",
			  currency = "15",
			  baseValue = "true",
			  annual = FALSE,
			  change = FALSE,
			  rank = FALSE) {
	# Setup
	require(quantmod)
	options(stringsAsFactors = FALSE)
	
	
	# load download code dataframe
	codeDF <- msci.list()
					  
	# extra codes given by user
	codes <- codeDF[codeDF$Country.Code %in% countries,]
	
	# Convert start and end dates to URL format
		# start date
		startdate <- as.Date(startDate)
		startchar <- format(startdate, format = "%d %b, %Y")
		startchar <- gsub(" ","%20",startchar)
		# end date
		enddate <- as.Date(endDate)
		endchar <- format(enddate, format = "%d %b, %Y")
		endchar <- gsub(" ","%20",endchar)
	
	# construct the URL
	URL <- paste0("https://www.msci.com/webapp/indexperf/charts?indices=",
					paste(codes$Download.Code, collapse = "|"),
					"&startDate=",
					startchar,
					"&endDate=",
					endchar,
					"&priceLevel=",
					priceLevel,
					"&currency=",
					currency,
					"&frequency=D&scope=R&format=XLS&baseValue=",
					baseValue,
					"&site=gimi")
	
	
	### Download and Process Initial Data ###
		# Download file
		download.file(URL, destfile = "msci file.xls", mode = "wb")
		# Load file into variable
		msciWB <- XLConnect::loadWorkbook("msci file.xls")
		# Extract sheet from WB
		index <- XLConnect::readWorksheet(msciWB, sheet = 1, startRow = 7)
		# Subset resulting DF to not include Copyright text at the bottom
		index <- index[1:(as.numeric(head(rownames(index[is.na(index$Date),]),1)) - 1), ]
	
		
	### Replace column names with user's codes ###
		# create named vector of country codes
		codesvec <- codes$Country.Code
		names(codesvec) <- codes$Col.Name
		# Replace index names with country codes by looking them up from named vector
		names(index) <- c("Date",codesvec[names(index)[-1]])
	
		
	### Further Processing ###
		#Convert to date
		index$Date <- as.Date(index$Date)
		# Fill in blanks
		if(length(countries)>1) {
			index[, -1] <- apply(index[, -1], 2, na.locf, na.rm = FALSE, maxgap = 12)
		} else {
			index[,2] <- na.locf(index[,2], na.rm = FALSE, maxgap = 12)
		}
		
	# Remove file after finishing
	file.remove("msci file.xls")
	
	
	#### Conditional Transformations ####
	
	## Convert to annual if requested
	if(annual) {
		index <- index[format(index$Date, format = "%m") == "12", ]
		index$Date <- format(index$Date, format = "%Y")
	}
	
	## Convert to % change if requested
	if(change) {
		if(length(countries) > 1) {
			index[, -1] <- apply(index[,-1], 2, Delt)
		} else {
			index[,2] <- as.numeric(Delt(index[,2]))
		}
	}
	
	## Convert to rank if requested
	if(rank) {
		# If more than 1 country, convert to rank
		if(length(countries) > 1) {
			index[, -1] <- t(apply(index[, -1], 1, function(x) rank(-x, na.last = "keep", ties.method = "min")))
		} else {
			# If only one country, return warning
			warning("Only one index downloaded, returning unranked value.")
		}
	}
	
	return(index)
}


# List countries with download codes
msci.list <- function() {
	codeDF <- read.csv("https://raw.githubusercontent.com/JeremyBowyer/MSCI-Indices/master/MSCI%20Download%20Codes.csv")
	return(codeDF)
}
