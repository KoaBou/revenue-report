function main() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var ordersSheet = spreadsheet.getSheetByName('orders');
    var incomeSheet = spreadsheet.getSheetByName('income');
  
    if (!ordersSheet || !incomeSheet){
      SpreadsheetApp.getUi().alert("Sheet 'orders' or 'income' is not exist. Please check sheet name");
      return;
    }
  
    var orders = processOrdersData(ordersSheet, incomeSheet);
    var summary = caculateSummary(orders);
  
    var summarySheet = spreadsheet.getSheetByName('summary');
    if (!summarySheet) {
      summarySheet = spreadsheet.insertSheet('summary');
    } else { 
      summarySheet.clear();
    }
    
    createStatistics(summary, summarySheet);
  }
  
  function processOrdersData(ordersSheet, incomeSheet) {
    // Get data from sheet
    var ordersData = ordersSheet.getDataRange().getValues();
    var ordersHeaders = ordersData[0];
    var ordersData = ordersData.slice(1);
  
    var incomeData = incomeSheet.getDataRange().getValues();
    var incomeHeaders = incomeData[0];
    var incomeData = incomeData.slice(1);
  
  
    
    // Get income indexes
    var incomeIdIndex = incomeHeaders.indexOf('Mã đơn hàng');
    var incomeTotalPriceIndex = incomeHeaders.indexOf('Giá gốc');
    var fixedFeeIndex = incomeHeaders.indexOf('Phí cố định');
    var serviceFeeIndex = incomeHeaders.indexOf('Phí Dịch Vụ');
    var paymentFeeIndex = incomeHeaders.indexOf('Phí thanh toán');
    var afiliateFeeIndex = incomeHeaders.indexOf('Phí hoa hồng Tiếp thị liên kết');
    var discountIndex = incomeHeaders.indexOf('Mã giảm giá');
    var deliveryFeeIndex = incomeHeaders.indexOf('Phí vận chuyển Người mua trả');
    var deliverySupportedIndex = incomeHeaders.indexOf('Phí vận chuyển được trợ giá từ Shopee');
    var deliveryRealFeeIndex = incomeHeaders.indexOf('Phí vận chuyển thực tế');
  
    // Ouput data
    var ordersList = {};
  
    // Process income data
    incomeData.forEach(function (row) {
      var orderId = row[incomeIdIndex];
  
      ordersList[orderId] = {};
      ordersList[orderId].orderTotalPrice = parseFloat(row[incomeTotalPriceIndex]);
      // ordersList[orderId].subsidy = 
      // ordersList[orderId].refund = 
      // ordersList[orderId].deliveryFee = 
      ordersList[orderId].afiliateFee = parseFloat(row[afiliateFeeIndex]);
      ordersList[orderId].fixedFee = parseFloat(row[fixedFeeIndex]);
      ordersList[orderId].serviceFee = parseFloat(row[serviceFeeIndex]);
      ordersList[orderId].paymentFee = parseFloat(row[paymentFeeIndex]);
      ordersList[orderId].discount = parseFloat(row[discountIndex]);
      ordersList[orderId].refundFee = parseFloat(row[deliveryFeeIndex]) + parseFloat(row[deliverySupportedIndex]) + parseFloat(row[deliveryRealFeeIndex]);
  
      ordersList[orderId].totalProductNum = 0;
      ordersList[orderId].totalReturnNum = 0;
      ordersList[orderId].products = {}
    }); 
  
  
    // Get order indexes
    var orderIdIndex = ordersHeaders.indexOf('Mã đơn hàng');
    var orderStatusIndex = ordersHeaders.indexOf('Trạng Thái Đơn Hàng');
    // var orderTotalPriceIndex = ordersHeaders.indexOf('Tổng giá trị đơn hàng (VND)');
    // var orderFixedFeeIndex = ordersHeaders.indexOf('Phí cố định');
    // var orderServiceFeeIndex = ordersHeaders.indexOf('Phí Dịch Vụ');
    // var orderPaymentFeeIndex = ordersHeaders.indexOf('Phí thanh toán');
    var productPriceIndex = ordersHeaders.indexOf('Giá gốc');
    var productNumIndex = ordersHeaders.indexOf('Số lượng');
    var productTotalPriceIndex = ordersHeaders.indexOf('Tổng giá bán (sản phẩm)')
    var orderSKUIndex = ordersHeaders.indexOf('SKU sản phẩm');
    var orderSKUTypeIndex = ordersHeaders.indexOf('SKU phân loại hàng')
    var productReturnNumIndex = ordersHeaders.indexOf('Số lượng sản phẩm được hoàn trả')
    var subsidyIndex = ordersHeaders.indexOf('Người bán trợ giá');
  
  
    ordersData.forEach(function (row) {
      var orderId = row[orderIdIndex];
      var sku = row[orderSKUIndex];
      var skuType = row[orderSKUTypeIndex];
  
  
      if (row[orderStatusIndex] != 'Hoàn thành') { return; }
      if (!ordersList[orderId]) {
        console.warn("Missing", orderId, "in income sheet!");
        return;
      }
  
      if (!ordersList[orderId].products[sku]){
        ordersList[orderId].products[sku] = {
          'totalProductNum': 0,
          'totalReturnNum': 0,
          'type': {
            [skuType]: {
              'productNum':  parseInt(row[productNumIndex]),
              'productReturnNum': parseInt(row[productReturnNumIndex]),
              'productTotalPrice': parseInt(row[productTotalPriceIndex]),
              'productPrice': parseInt(row[productPriceIndex]),
              'subsidy': parseFloat(row[subsidyIndex]),
            }
          }
        }
      } else {
        ordersList[orderId].products[sku].type[skuType] = {
        'productNum':  parseInt(row[productNumIndex]),
        'productReturnNum': parseInt(row[productReturnNumIndex]),
        'productTotalPrice': parseInt(row[productTotalPriceIndex]),
        'productPrice': parseInt(row[productPriceIndex]),
        'subsidy': parseFloat(row[subsidyIndex]),
        };
      }
  
      ordersList[orderId].totalProductNum += parseInt(row[productNumIndex]);
      ordersList[orderId].totalReturnNum += parseInt(row[productReturnNumIndex]);
      ordersList[orderId].products[sku].totalProductNum += parseInt(row[productNumIndex]);
      ordersList[orderId].products[sku].totalReturnNum += parseInt(row[productReturnNumIndex]);
    });
  
    return ordersList;
  }
  
  
  function caculateSummary(orders) {
    // Sumary variables
    var summary = {};
  
    for (let orderID in orders){
      var order = orders[orderID];
  
      for (let sku in order.products){
        if (order.totalProductNum > 0){
          if ((order.totalProductNum - order.totalReturnNum) > 0) {
            var contribute = (order.products[sku].totalProductNum - order.products[sku].totalReturnNum) / (order.totalProductNum - order.totalReturnNum); 
          }
         if (order.totalReturnNum > 0){
          var contributeReturn = order.products[sku].totalReturnNum / order.totalReturnNum; 
         }
        }
  
        if (!summary[sku]){
          summary[sku] = {
            'revenue': 0,
            'orderNum': 0,
            'orderReturnNum': 0,
            'orderCompletedNum': 0,
            'fixedFee': 0,
            'serviceFee': 0,
            'paymentFee': 0,
            'afiliateFee': 0,
            'discount': 0,
            'subsidy': 0,
            'refund': 0,
            'refundFee': 0,
          }
        }
  
        var isReturn = false;
  
        for (let skuType in order.products[sku].type) {
          if (!isReturn) {
            if (order.products[sku].type[skuType].productReturnNum > 0) {
              isReturn = true;
            }
          }
          summary[sku].revenue += (order.products[sku].type[skuType].productPrice * (order.products[sku].type[skuType].productNum));
          summary[sku].refund += order.products[sku].type[skuType].productPrice * order.products[sku].type[skuType].productReturnNum;
          summary[sku].subsidy += order.products[sku].type[skuType].subsidy;
        }
        if (contribute){
          summary[sku].fixedFee += order.fixedFee * contribute;
          summary[sku].serviceFee += order.serviceFee * contribute;
          summary[sku].paymentFee += order.paymentFee * contribute;
          summary[sku].afiliateFee += order.afiliateFee * contribute;
          summary[sku].discount += order.discount * contribute;
        }
  
        if (contributeReturn) {
          summary[sku].refundFee += order.refundFee * contributeReturn;
        }
  
        if (isReturn) {
          summary[sku].orderReturnNum += 1;
        } else {
          summary[sku].orderCompletedNum += 1;
        }
  
        summary[sku].orderNum += 1;
      }
    }
  
    return summary;
  }
  
  function createStatistics(summary, summarySheet) {
    var headers = ['NỘI DUNG', 'DOANH THU', 'Doanh thu thuần', 'Đơn hàng', 'Đơn hoàn thành', 'Đơn trả hàng ', 'CHI PHÍ', 'Phát sinh hoàn', 'Phí cố định', 'Phí dịch vụ', 'Phí thanh toán', 'Tiếp thị liên kết', 'Mã giảm giá', 'Trợ giá sản phẩm', 'Hoàn hàng'];
  
    var rows = [];
  
    for (let sku in summary){
      var row = [sku, '', summary[sku].revenue, summary[sku].orderNum, summary[sku].orderCompletedNum, summary[sku].orderReturnNum, '', -summary[sku].refundFee, -summary[sku].fixedFee, -summary[sku].serviceFee, -summary[sku].paymentFee, -summary[sku].afiliateFee, -summary[sku].discount, summary[sku].subsidy, summary[sku].refund];
  
      rows.push(row);
    }
  
    for (let i = 0; i < headers.length; i++){
      var row = [];
      row.push(headers[i]);
  
      for (let j=0; j < rows.length; j++){
        row.push(rows[j][i]);
      }
  
      summarySheet.appendRow(row);
    }
  }