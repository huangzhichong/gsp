# encoding: UTF-8

require 'spreadsheet'
require_relative 'database'
require 'pry'
require 'pry-debugger'

def get_source_data(file_name = 'TMDC.xls')
  Spreadsheet.client_encoding = 'UTF-8'
  doc = Spreadsheet.open file_name
  tmdc_sheet = doc.worksheet 'TMDC'
  headers = tmdc_sheet.row(0)
  column_counts = headers.count
  result  = []
  tmdc_sheet.each do |row|
    temp = Hash.new
    column_counts.times do |i|
      temp[headers[i-1]] = row[i-1]
    end
    result << temp
  end
  result.delete_at(0)
  return result
end

def timestamp
  sleep 1
  Time.now.strftime('%Y%m%d%H%M%S')
end

get_source_data.each do |data|

  po_number ="P#{timestamp}"
  #make sure there is an record for GoodsDefinition
  #TODO
  # 联系人，联系方式，生产商，供应商商如何获取
  # 原始数据包含规格和型号，合并入型号栏？
  data['产品注册号'] = "#{data['产品注册证']} #{data['产品注册证号']}"
  t = $sqlserver_env_util.query "select * from GoodsDefinition where 备注 = '#{data['商品编码']}'"
  unless t.length > 0
    sql = "INSERT INTO GoodsDefinition (商品名称,型号,基本单位,生产厂商,供应商,产品注册号,联系方式,联系人,备注,修改人,修改时间)
           VALUES ('#{data['商品名称']}','#{data['型号']}','#{data['单位']}','#{data['产地']}','#{data['产地']}','#{data['产品注册号']}','8008700302','联系人','#{data['商品编码']}','Administrator','#{Time.now.strftime('%Y-%m-%d')}');"
    $sqlserver_env_util.query sql
  end



  t = $sqlserver_env_util.query "select top 1 商品编号 from GoodsDefinition where 备注 = '#{data['商品编码']}'"
  data['商品编号'] = t[0]['商品编号']

  #add purchase order info
  #TODO
  # 订单号自动按时间生成
  # 供应商同GoodsDefinition
  # 采购员？
  # 到货时间，订单日期如何计算
  sql = "INSERT INTO PurchaseOrder_main(订单号,供应商,采购员,预计到货日期,摘要,附加说明,订单日期,是否完成验收,修改人,修改时间)
         VALUES ('#{po_number}','#{data['产地']}','采购员','#{Time.now.strftime('%Y-%m-%d')}','无','无','#{Time.now.strftime('%Y-%m-%d')}',0,'Administrator','#{Time.now.strftime('%Y-%m-%d')}')"
  $sqlserver_env_util.query sql

  #add purchase order details
  #TODO
  #具体信息同商品信息
  sql = "INSERT INTO PurchaseOrder_detail(订单号,商品编号,商品名称,型号,基本单位,订单数量,生产厂商,供应商,产品注册号,联系方式,联系人,备注)
         VALUES ('#{po_number}', '#{ data['商品编号']}','#{data['商品名称']}','#{data['型号']}','#{data['单位']}','#{data['数量']}','#{data['产地']}','#{data['产地']}','#{data['产品注册号']}','8008700302','联系人','无')"
  $sqlserver_env_util.query sql


end

puts $sqlserver_env_util.query "select * from PurchaseOrder_main"
