require 'rubygems' # To use all the Gems
require 'nokogiri' # To collect H1s and H2s
require 'open-uri' # Goes with Nokogiri
require 'metainspector' # To collect the meta-infos
require 'roo' # To handle Excel reading
require 'csv' # To create the Final CSV
require 'watir' # To take the screenshot


# Global int -> Keeps track of the number of errors
@count = 0


# Aim of the function :
# Create the proper header
# Then add it to the CSV File

def 		add_header_to_csv(file)
	row = []
	
	# Add the first row in the CSV : all the Arguments
	row << "URL"
	row << "META-TITLE"
	row << "META-DESCRIPTION"
	row << "META-KEYWORDS"
	row << "H1s"
	row << "H2s"
	row << "Name_screenshot"

	# Push the row to the CSV
	file << row
end


# Aim of the function :
# Scrape all the metadatas
# Using the Metainspector gem

def take_screenshot(browser, url, row, number_image)

	# go to the URL, but if it takes too long, abort
	Watir::Wait.until(20) { browser.goto url }
	# browser.goto url
	browser.screenshot.save "#{number_image}.png"
	row << number_image
	
	return row
end


# Aim of the function :
# Scrape all the metadatas
# Using the Metainspector gem

def meta_inspector_scrapping(row, url)
	
	# Open the page
	if (page = MetaInspector.new(url, :encoding => 'UTF-8'))

		# Collect META-TITLE
		if (page.title)
			row << page.title
		else
			row << "No Title"
		end

		# Collect META-DESCRIPTION
		if (page.description)
			row << page.description
		else
			row << "No Description"
		end

		# Collect KEYWORDS
		if (page.meta_tag['name']['keywords'])
			row << page.meta_tag['name']['keywords']
		else
			row << "No Keyword"
		end

	# If it didn't work out, print these messages
	else
		row << "No collect possible"
		row << "No collect possible"
		row << "No collect possible"
	end

	return row
end


# Aim of the function :
# Use of the second Gem : Nokogiri
# Nokogiri will collect all the h1s and h2s

def nokogiri_scrapping(row, url)

	# First we open the page with Nokogiri
	if (doc = Nokogiri::HTML(open(url)))

		# Collect H1s
			if doc.at('h1')
				row << doc.at('h1')
			else
				row << "No h1"
			end

		# Collect H2s
			if doc.at('h2')
				row << doc.at('h2')
			else
				row << "No h2"
			end

	else
		row << "No collect-2 possible"
		row << "No collect-2 possible"
	end

	return row
end


# Aim of the function :
# Heart of the scrapping part of the program
# We escape any error to ensure that the program will parse
# the maximum amount of URLs
# We first clear the array, then we add all the infos to the "row" array
# Array that is filled in a specific order to match the columns of the CSV
# At the end, we return the array, which is the row that will be added to the CSV

def heart_of_scraping(row, url, browser, number_image)
	# Error management, in order to avoid it to stop the program
	begin
		
			# First, clear the array
			row.clear

			# Collect URL
			row << url

			# Launches the first part of the scraping (metadatas)
			row = meta_inspector_scrapping(row, url)
			# Launches the second part of the scraping (h1s and h2s)
			row = nokogiri_scrapping(row, url)
			# Launches the third part of the scraping (the screenshot + name of the file)
			row = take_screenshot(browser, url, row, number_image)

	# Error management, we catch everything to be sure that the program doesn't stop
	rescue => e
		# To keep track of the number of URLs that didn't work out 
		@count = @count + 1
	end

	return row
end


# Aim of the function :
# Manage all the scrapping functions
# Opens and create the CSV FILE
# with the name "outcome.csv"
# Then adds the header to the csv
# Then loops on each URL and scrapes the data

def 		get_all_the_infos(urls)
	
	number_image = 0;
	
	puts "We collected your URLs, we will now scrape them"
	# urls.clear
	# urls = ["https://www.skyscrapercity.com/showthread.php?t=670412","http://bridgestone.cfaomotors-togo.com/fr/catalog/produits/4x4", "http://www.uae-business-directory.com/directory/dubai/sheikh-zayed-road/tyre-dealers/bridgestone.html", "http://www.seodirectory6.com/Shopping/Auctions/page-2.html"]
	
	# Variable that will stock all the infos for the CSV
	row = []

	# To bypass https / http certificate problems
	capabilities = Selenium::WebDriver::Remote::Capabilities.firefox(accept_insecure_certs: true)
	# client = Selenium::WebDriver::Remote::Http::Default.new
	# client.timeout = 30 # seconds – default is 60

	# Opens a browser (Firefox in this case)
	# See the actions of the browser ->
	browser = Watir::Browser.new(:firefox, :desired_capabilities => capabilities)
	# Headless mode (everything happens in background) -> Uncomment the next line, and comment the upper one
	# browser = Watir::Browser.new(:firefox, :desired_capabilities => capabilities, :headless => true) # , :headless => true || , :http_client => client

	# Create a new EXCEL FILE that will store all the data
	CSV.open("outcome.csv", "w") do |file|
		
		# Add the header to the csv file
		add_header_to_csv(file)
		
		# Iterate on each url in order to collect the proper info
		urls.each do | url |
		
			row = heart_of_scraping(row, url, browser, number_image)
			# Add the row to the CSV
			file << row
			number_image = number_image + 1
		
		end
		
		# prints the number of errors
		puts "Nombre d'URL non parsées: #{@count}"
	
	end
	
	# Closes the browser
	browser.close
	puts "Finished"
end


# Aim of the function :
# Function that initializes the program
# This function opens the excel
# Then reads the first column
# And stocks all the urls inside the links variable
# Once done, it passes all the infos
# To the get_all_the_infos function

def 	read_from_csv()

	links = []
	
	# open excel file
	xlsx = Roo::Spreadsheet.open('./fichier_source.xlsx')
	
	# select the proper sheet. The first one is the basis
	feuille = xlsx.sheets[0]

	# Collect all the URLs from the sheet, column A
	links = xlsx.sheet(feuille).column(1)

	# Removes the first element of the column (the name of the table)
	links = links.drop(1)

	# Send all the info to the get_all_the_infos function that
	# will parse all the infos
	get_all_the_infos(links)

end

read_from_csv()