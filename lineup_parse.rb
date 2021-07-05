require 'docx'
require 'csv'
# replace "FILE-NAME.docx" with the file name
@file = Docx::Document.open('./Files/Schedules/FILE-NAME.docx')

@months = ["January", "Feb", "March", "April", "May", "June", "July", "Aug", "Sept", "Oct", "Nov", "Dec"]

@raw_text = []
@nested_text = {}
@result = {}

# grabs the raw text + formatting line by line (excluding spaces and time, which have no use here) and stores it in an array
def initial_parse(file, raw_text)
    file.paragraphs.each do |paragraph|
        raw_text << paragraph unless paragraph.text.empty? || time?(paragraph)
    end
end

# takes the raw_text and organizes it into a nested structure of date => project => [worker names]
def text_sort(raw_text, nested_text)
    current_date = ""
    current_worker = ""
    current_project = ""
    raw_text.each do |line|
        next if line.text.include?("Schedule")
        clean_string = line.text.strip
        if date?(line)
            current_date = clean_string
            nested_text[current_date] = {}
        elsif project_name?(line)
            current_project = clean_string
            nested_text[current_date][current_project] = []
        elsif worker_name?(line) || worker_name_with_note?(line)
            # p clean_string
            # p current_date
            # p current_project
            # p nested_text[current_date][current_project]
            nested_text[current_date][current_project] << clean_string
        end
    end
end

def sort_workers(nested_text, result)
    current_date = ""
    current_project = ""
    current_worker = ""
    #goes through the nested text...
    nested_text.each do |date, projects|
        #first assigns the current date...
        current_date = date.to_s
        projects.each do |project, workers|
            #then looks in the projects hash value and assigns the current project...
            current_project = project.to_s
            workers.each do |worker|
                #then looks in the workers hash value and assigns the current worker...
                current_worker = worker.to_s
                #then creates a worker => date => [projects] nest in the result hash
                result[current_worker] = {} unless result[current_worker]
                result[current_worker][current_date] = [] unless result[current_worker][current_date]
                result[current_worker][current_date] << current_project
            end
        end
    end
end

#if the paragraph starts with a number, it's the time
def time?(paragraph)
    return paragraph.text[0].match(/\A\d+\z/)
end

#if the paragraph contains a month, it's a date
def date?(paragraph)
    result = false
    @months.each do |month|
        result = true if paragraph.text.include?(month) && !paragraph.text.include?("Schedule")
    end
    result
end

#if the paragraph contains an emphasized character, it's a project name
def project_name?(paragraph)
    return paragraph.to_html.include?("strong") && (!worker_name?(paragraph) && !date?(paragraph))
end

def header?(paragraph)
    return paragraph.text.include?("Schedule")
end

#if the paragraph contains between one and two words and no digits, it's a worker's name
def worker_name?(paragraph)
    return paragraph.text.split(" ").length.between?(1, 2) && (!paragraph.text.match?(/\d/) && !paragraph.text.include?("Belmont"))
end

def worker_name_with_note?(paragraph)
    return !project_name?(paragraph) && paragraph.text.include?("(")
end

#creates the CSV file
def create_sheets(result)
    current_worker = ""
    current_date = ""
    current_project = ""
    result.each do |worker, nest|
        current_worker = worker.to_s
        path = "./Files/Finished Files/#{current_worker}.csv"
        nest.each do |date, projects|
            current_date = date.to_s
            projects.each do |project|
                current_project = project.to_s
                p path
                CSV.open(path, "a") do |csv|
                    #we only need to add the worker's name to the file if it is empty
                    csv << [current_worker] if CSV.readlines(path).take(1).empty?
                    csv << [current_date, current_project]
                end
            end
        end
        #adds an empty line (makes the final document easier to read as we have to merge the files via command line)
        CSV.open(path, "a") { |csv| csv << [nil] }
    end
end

def driver
    initial_parse(@file, @raw_text)
    text_sort(@raw_text, @nested_text)
    sort_workers(@nested_text, @result)
    create_sheets(@result)
end

driver

