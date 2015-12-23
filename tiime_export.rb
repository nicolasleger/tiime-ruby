require 'json'
require 'mechanize'
require 'net/http'
require 'nokogiri'
require 'open-uri'

DIRECTORY = File.expand_path(File.dirname(__FILE__))
EXPORTS_DIRECTORY_NAME = File.join(DIRECTORY, "exports")

TIIME_HOST = "https://secure.tiime.fr"

def balance(agent)
  dom = Nokogiri::HTML(agent.page.body)
  balance_element = dom.css('#montants > .bloc_solde > span').first
  balance = balance_element.text[/\d[\d ,€]+/]
end

def login(email, password, agent, new_version = false)
  url = new_version ? 'https://new.tiime.fr' : TIIME_HOST
  page = agent.get(url)

  login_form = page.forms.first
  login_form.usr = email
  login_form.pwd = password

  page = login_form.submit

  # Redirection to new login form
  new_login_form = page.forms.first
  if !new_version && !new_login_form.nil? && new_login_form.action.include?('new.tiime.fr')
    return login(email, password, agent, true)
  end
  
  page.uri.to_s.include?('banques')
end

def get_document_ids(agent)
  # 1. Browse master folder
  page = agent.post("#{TIIME_HOST}/documents.requests.php", { queryname: 'chargeexplorer' })
  folder = JSON.parse(page.body)

  # 2. Browse year subfolders
  folder_ids = (2008..Date.today.year).to_a.reverse.flat_map { |year|
    subfolders_tree = folder['folder'][year.to_s]['ArborescenceDocument']

    subfolders_tree_dom = Nokogiri::HTML(subfolders_tree)
    subfolders_tree_dom.css('li').map { |subfolder|
      subfolder.attr('data-hdl')
    }
  } # Recent to old

  # 3. Browse folders
  document_ids = folder_ids.flat_map do |folder_id|
    page = agent.post("#{TIIME_HOST}/documents.requests.php", {
      queryname: 'chargeviewer',
      handle: folder_id,
      search: ''
    })

    documents_tree = JSON.parse(page.body)
    documents_tree_dom = Nokogiri::HTML(documents_tree['liste'])
    documents_tree_dom.css('tbody tr').map { |subfolder|
      subfolder.attr('data-hdl')
    }
  end.compact

  document_ids
end

def mkdir_exports_directory
  Dir.mkdir(EXPORTS_DIRECTORY_NAME) unless File.exists?(EXPORTS_DIRECTORY_NAME)
end

def download_bank_report(agent)
  # 1. Need to load export
  agent.get("#{TIIME_HOST}/banques.php?filtre=")
  agent.post("#{TIIME_HOST}/banques.requests.php", {
    queryname: 'charge',
    pagecourante: '1',
    senspage: 'init',
    senstri: 'DESC',
    colonnetri: 'date',

    filtreimputations: '',
    filtreimputationspecial: '',
    filtrepjs: '',
    searchlabel: '',
    filtredates: '',
    filtretypesoperations: '',
    filtrebanques: ''
  })

  # 2. Download
  mkdir_exports_directory
  filename = "bank_report.xls"
  filepath = File.join(EXPORTS_DIRECTORY_NAME, filename)
  puts "Download bank report to #{filepath}"

  file_content = agent.get("#{TIIME_HOST}/export.php").body
  return false if file_content.to_s == "Une erreur s'est produite, export impossible !"

  open(filepath, 'wb') { |file|
    file << file_content
  }

  true
end

def download_personal_fees(agent)
  # 1. Need to load export
  agent.post("#{TIIME_HOST}/depense.requests.php?filtre=", {
    queryname: 'charge',
    senspage: 'init',
    senstri: 'DESC',
    colonnetri: 'date',
    filtreimputations: '',
    filtreimputationspecial: '',
    searchlabel: '',
    filtredates: '',
    filtrepjs: '',
    pagecourante: '1'
  })

  # 2. Download
  mkdir_exports_directory
  filename = "personal_fees.xls"
  filepath = File.join(EXPORTS_DIRECTORY_NAME, filename)
  puts "Download personal fees to #{filepath}"

  file_content = agent.get("#{TIIME_HOST}/export.php", [], "#{TIIME_HOST}/depense.php").body
  return false if file_content.to_s == "Une erreur s'est produite, export impossible !"

  open(filepath, 'wb') { |file|
    file << file_content
  }

  true
end

def download_documents(agent)
  documents_directory_name = 'documents'
  Dir.mkdir(documents_directory_name) unless File.exists?(documents_directory_name)

  document_ids = get_document_ids(agent)
  puts "Download documents"
  document_ids.each do |document_id|
    # 1. Need to load document before
    agent.post("#{TIIME_HOST}/documents.requests.php", {
      queryname: 'voirdocument',
      handle: document_id,
      viewerwidth: '-30'
    })

    # 2. Download
    filename = "#{document_id}.pdf"
    filepath = File.join(EXPORTS_DIRECTORY_NAME, filename)
    puts "Download #{filename} to #{filepath}"
    open(filepath, 'wb') { |file|
      # No need to be logged, each file is publicly available! (but after being charged just before)
      file << open("#{TIIME_HOST}/view/#{filename}").read
    }
  end

  true
end

def run!(email, password)
  agent = Mechanize.new
  
  # 1. Login
  if login(email, password, agent)
    puts "Logged"

    puts "Balance #{balance(agent)}"

    # 2. Download banking export
    download_bank_report(agent)

    # 3. Download persional fees export
    download_personal_fees(agent)

    # 4. Download documents
    download_documents(agent)
  else
    raise "Not logged successfully"
  end
end

if __FILE__ == $0
  if ARGV.length < 2
    puts "Usage: ruby tiime_export.rb EMAIL PASSWORD"
  else
    run!(ARGV[0], ARGV[1])
  end
end