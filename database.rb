require 'tiny_tds'
require 'yaml'

class DatabaseUtil
  def initialize
    @envHash = Hash.new
    @envHash[:mode]="dblib"
    @envHash[:timeout]=5000
    @envHash[:username]='sa'
    @envHash[:password]='@ctive123'
    @envHash[:dataserver]='10.109.0.161'
    @envHash[:database]= 'SMARTTEST'
  end
  def query(sql)
    result = []
    client = TinyTds::Client.new(@envHash)
    begin
      result = client.execute(sql).to_a
      return result
    rescue => e
      raise "error in execute_query -- #{sql} \n #{e.message}"
    end
  ensure client.close unless (client.nil? || client.closed?)
  end

  def query_return_affected_rows(sql)

    result = []
    client = TinyTds::Client.new(@envHash)
    begin
      result = client.execute(sql)
      rows = result.affected_rows
      return rows
    rescue => e
      raise "error in execute_query -- #{db_category+sql} \n #{e.message}"
    end
  ensure client.close unless (client.nil? || client.closed?)
  end
end
$sqlserver_env_util = DatabaseUtil.new