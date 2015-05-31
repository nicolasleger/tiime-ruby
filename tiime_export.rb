require 'json'
require 'mechanize'
require 'net/http'
require 'nokogiri'
require 'open-uri'

def login(email, password, agent)
  page = agent.get('https://secure.tiime.fr')

  form = page.forms.first
  form['usr'] = email
  form['pwd'] = password
  form['rco'] = 'on' # Stay connected

  page = form.submit
  
  !page.title.include?('Connectez-vous')
end

def get_document_ids(agent)
  # 1. Browse master folder
  params = { queryname: 'chargeexplorer' }
  page = agent.post('https://secure.tiime.fr/documents.requests.php', params)
  folder = JSON.parse(page.body)

  # 2. Browse year subfolders
  folder_ids = (2012..2015).map { |year|
    subfolders_tree = folder['folder'][year.to_s]['ArborescenceDocument']

    subfolders_tree_dom = Nokogiri::HTML(subfolders_tree)
    subfolders_tree_dom.css('li').map { |subfolder|
      subfolder.attr('data-hdl')
    }
  }.reverse.flatten # Recent to old

  # 3. Browse folders
  document_ids = folder_ids.map do |folder_id|
    params = {
      queryname: 'chargeviewer',
      handle: folder_id,
      search: ''
    }
    page = agent.post('https://secure.tiime.fr/documents.requests.php', params)

    documents_tree = JSON.parse(page.body)
    documents_tree_dom = Nokogiri::HTML(documents_tree['liste'])
    documents_tree_dom.css('tbody tr').map { |subfolder|
      subfolder.attr('data-hdl')
    }
  end.flatten.compact

  document_ids
end

def download_bank_report(agent)
  exports_directory_name = 'exports'
  Dir.mkdir(exports_directory_name) unless File.exists?(exports_directory_name)

  filename = "bank_report.xls"

  # 1. Need to load export
  agent.post('https://secure.tiime.fr/banques.requests.php', {
    queryname: 'charge',
    senspage: 'init',
    senstri: 'DESC',
    colonnetri: 'date',
    filtreimputations: '',
    filtreimputationspecial: '',
    searchlabel: '',
    filtredates: '',
    filtrepjs: '',
    pagecourante: '1',

    filtretypesoperations: '',
    filtrebanques: ''
  })

  # 2. Download
  puts "Download bank report"
  file_content = agent.get('https://secure.tiime.fr/export.php').body
  return false if file_content.to_s == "Une erreur s'est produite, export impossible !"

  open("#{exports_directory_name}/#{filename}", 'wb') { |file|
    file << file_content
  }

  true
end

def download_personal_fees(agent)
  exports_directory_name = 'exports'
  Dir.mkdir(exports_directory_name) unless File.exists?(exports_directory_name)

  filename = "personal_fees.xls"
  
  # 1. Need to load export
  agent.post('https://secure.tiime.fr/depense.requests.php', {
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
  puts "Download personal fees"
  file_content = agent.get('https://secure.tiime.fr/export.php').body
  return false if file_content.to_s == "Une erreur s'est produite, export impossible !"

  open("#{exports_directory_name}/#{filename}", 'wb') { |file|
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
    params = {
      queryname: 'voirdocument',
      handle: document_id,
      viewerwidth: '-30'
    }
    agent.post('https://secure.tiime.fr/documents.requests.php', params)

    # 2. Download
    filename = "#{document_id}.pdf"
    puts "Download #{filename}"
    open("#{documents_directory_name}/#{filename}", 'wb') { |file|
      # No need to be logged, each file is publicly available! (but after being charged just before)
      file << open("https://secure.tiime.fr/view/#{filename}").read
    }
  end

  true
end

def run!(email, password)
  agent = Mechanize.new
  
  # 1. Login
  if login(email, password, agent)
    puts "Logged"

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