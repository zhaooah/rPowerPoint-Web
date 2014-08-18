require "bundler/setup"
Bundler.require(:default)
require './lib/rPowerPoint'
 
 configure :production, :development do
   enable :logging
   set :sessions, true
set :logging, true
set :dump_errors, false
set :some_custom_option, false
 end

before { env['rack.logger'] = Logger.new'testlog'  }

def GeneratePowerPoint(outputName)

  #Read Input CSV
  src={}
  CSV.foreach('./'+outputName+'/dataource.csv') do |row|
    temp=[]
    row.drop(1).each do |c|
      if c.nil?
      else
        src[row[0]]=temp.push(c)
      end
    end
  end

  #Call the function
  RPowerPoint::PowerPointObject.new(outputName,src)

  haml :result


end

get '/' do
    haml :index
end

post '/' do
    FileUtils::mkdir_p params['outputName']

    File.open(params['outputName']+ '/template.pptx', "w") do |f|
      f.write(params['template'][:tempfile].read)
    end

    File.open(params['outputName']+ '/mapping.csv', "w") do |f|
      f.write(params['mapping'][:tempfile].read)
    end

    File.open(params['outputName']+ '/dataource.csv', "w") do |f|
      f.write(params['datasource'][:tempfile].read)
    end

    GeneratePowerPoint(params['outputName'])

end
