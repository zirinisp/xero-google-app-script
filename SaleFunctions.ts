
var AllAccountingDataTable = 'All Accounting Data';
var SalesDataTable = 'Day Sales';


namespace Accounting {
    enum StatusEnum {
        Draft = 'DRAFT',
        Authorised = 'AUTHORISED',
        Paid = 'PAID',
        Deleted = 'DELETED',
        Voided = 'VOIDED',
        Other = ''
    }

    class Status {
        status: string;
        constructor(status: string) {
            this.status = status;
        }

        active(): Boolean {
            if ((this.status === StatusEnum.Paid) || (this.status === StatusEnum.Authorised)) {
                return true;
            }
            return false;
        }
    }

    enum AccountEnum {
        Sales = '200',
        DeliverySales = '201',
        UberEats = 'UBER',
        Deliveroo = 'DELR',
        MealPal = 'MPAL',
        Seamless = 'SEAML',
        Cash = 'UND',
        iZettle = 'IZET',
        FoodCost = '310',
        Contructors = '481',
        Other = '',
        KSeries = 'K-CLR'
    }

    class Account {
        account: string;
        description: string;
        constructor(account: string, description: string) {
            this.account = account;
            this.description = description;
        }

        isSales(): Boolean {
            if ((this.account.toString() === AccountEnum.Sales) || (this.account.toString() === AccountEnum.DeliverySales)) {
                return true;
            }
            return false;
        }

        isRestaurant(): Boolean {
            if (this.account.toString() === AccountEnum.Sales) {
                return true;
            }
            return false;
        }

        isDelivery(): Boolean {
            if (this.account.toString() === AccountEnum.DeliverySales) {
                return true;
            }
            return false;
        }

        isUber(): Boolean {
            if (this.account.toString() === AccountEnum.UberEats || (this.account.toString() === AccountEnum.KSeries && this.description.indexOf('UberEats') !== -1)) {
                return true;
            }
            return false;
        }

        isDeliveroo(): Boolean {
            if (this.account.toString() === AccountEnum.Deliveroo || (this.account.toString() === AccountEnum.KSeries && this.description.indexOf('Deliveroo') !== -1)) {
                return true;
            }
            return false;
        }
        isCash(): Boolean {
            if (this.account.toString() === AccountEnum.Cash || (this.account.toString() === AccountEnum.KSeries && this.description.indexOf('Cash') !== -1)) {
                return true;
            }
            return false;
        }
        isCard(): Boolean {
            if (this.account.toString() === AccountEnum.iZettle || (this.account.toString() === AccountEnum.KSeries && this.description.indexOf('iZettle') !== -1)) {
                return true;
            }
            return false;
        }
        isSeamless(): Boolean {
            if (this.account.toString() === AccountEnum.Seamless || (this.account.toString() === AccountEnum.KSeries && this.description.indexOf('Seamless') !== -1)) {
                return true;
            }
            return false;
        }
        isMealPal(): Boolean {
            if (this.account.toString() === AccountEnum.MealPal || (this.account.toString() === AccountEnum.KSeries && this.description.indexOf('MealPal') !== -1)) {
                return true;
            }
            return false;
        }
        isFoodCost(): Boolean {
            if (this.account.toString() === AccountEnum.FoodCost) {
                return true;
            }
            return false;
        }
    }

    enum TrackingCategoryEnum {
        Delivery_Bakaliko = 'DEL-Bakaliko',
        Delivery_Grill = 'DEL-Grill',
        Delivery_Salads = 'DEL-Salads',
        Delivery_Kitchen = 'DEL-Kitchen',
        Delivery_SoftDrinks = 'DEL-SoftDrinks',
        Delivery_Desserts = 'DEL-Desserts',
        Delivery_Wines = 'DEL-Wines',
        Delivery_Deliverect_Both = 'DEL-DeliverectBoth',
        Delivery_Beer_Shots = 'DEL-Beer-Shots',
        Restaurant_Salads = 'RES-Salads',
        Restaurant_Wines = 'RES-Wines',
        Restaurant_Grill = 'RES-Grill',
        Restaurant_Kitchen = 'RES-Kitchen',
        Restaurant_Various = 'RES-Various',
        Restaurant_Beer_Shots = 'RES-Beer-Shots',
        Restaurant_Bakaliko = 'RES-Bakaliko',
        Restaurant_SoftDrinks = 'RES-SoftDrinks',
        Restaurant_Desserts = 'RES-Desserts',
        Restaurant_Coffee = 'RES-Coffee',
        Pantelis = 'PZ',
        Argyris = 'AR',
        Kostis = 'KM',
        Delivery_Zero_Rated = 'DLV0',
        Other = ''
    }

    class TrackingCategory {
        category: string;
        constructor(category: string) {
            this.category = category;
        }

        isKitchen(): Boolean {
            if (
                (this.category.toString() === TrackingCategoryEnum.Delivery_Salads) ||
                (this.category.toString() === TrackingCategoryEnum.Restaurant_Salads) ||
                (this.category.toString() === TrackingCategoryEnum.Delivery_Kitchen) ||
                (this.category.toString() === TrackingCategoryEnum.Restaurant_Kitchen) ||
                (this.category.toString() === TrackingCategoryEnum.Delivery_Desserts) ||
                (this.category.toString() === TrackingCategoryEnum.Restaurant_Desserts)) {

                return true;
            }
            return false;
        }

        isGrill(): Boolean {
            if (
                (this.category.toString() === TrackingCategoryEnum.Delivery_Grill) ||
                (this.category.toString() === TrackingCategoryEnum.Restaurant_Grill)) {
                return true;
            }
            return false;
        }

        isService(): Boolean {
            if (
                (this.category.toString() === TrackingCategoryEnum.Delivery_Bakaliko) ||
                (this.category.toString() === TrackingCategoryEnum.Restaurant_Bakaliko) ||
                (this.category.toString() === TrackingCategoryEnum.Delivery_Wines) ||
                (this.category.toString() === TrackingCategoryEnum.Restaurant_Wines) ||
                (this.category.toString() === TrackingCategoryEnum.Delivery_SoftDrinks) ||
                (this.category.toString() === TrackingCategoryEnum.Restaurant_SoftDrinks) ||
                (this.category.toString() === TrackingCategoryEnum.Restaurant_Beer_Shots) ||
                (this.category.toString() === TrackingCategoryEnum.Restaurant_Coffee) ||
                (this.category.toString() === TrackingCategoryEnum.Delivery_Beer_Shots)) {
                return true;
            }
            return false;
        }


    }

    enum EntryTypeEnum {
        ACCPAY = 'ACCPAY',
        ACCREC = 'ACCREC',
        RECEIVE_TRANSFER = 'RECEIVE-TRANSFER',
        SPEND_TRANSFER = 'SPEND-TRANSFER',
        RECEIVE = 'RECEIVE',
        SPEND = 'SPEND',
        SPEND_PREPAYMENT = 'SPEND-PREPAYMENT',
        RECEIVE_PREPAYMENT = 'RECEIVE-PREPAYMENT',
        SPEND_OVERPAYMENT = 'SPEND-OVERPAYMENT',
        RECEIVE_OVERPAYMENT = 'RECEIVE-OVERPAYMENT'
    }

    class EntryType {
        type: string;
        constructor(type: string) {
            this.type = type;
        }
        // The following is not correct we need to find account information to get these data.
        isExpense(): Boolean {
            if (
                (this.type.toString() === EntryTypeEnum.ACCREC) ||
                (this.type.toString() === EntryTypeEnum.ACCPAY) ||
                (this.type.toString() === EntryTypeEnum.SPEND)) {

                return true;
            }
            return false;
        }

/* Not finished
        isPrepayment() : Boolean {
            if (
                (this.type.toString() === EntryTypeEnum.SPEND_PREPAYMENT) ||
                (this.type.toString() === EntryTypeEnum.RECEIVE_PREPAYMENT) ||
                (this.type.toString() === EntryTypeEnum.SPEND_OVERPAYMENT) ||
                (this.type.toString() === EntryTypeEnum.RECEIVE_OVERPAYMENT) ||
                (this.type.toString() === EntryTypeEnum.RECEIVE_TRANSFER) ||
                (this.type.toString() === EntryTypeEnum.SPEND_TRANSFER)) {

                return true;
            }
            return false;

        }*/
    }

    export class Sales {
        entries: Entry[];
        constructor(entries: Entry[]) {
            this.entries = entries;
        }

        salesForPeriod(from: Date, to: Date): Sales {
            var newEntries: Entry[] = [];
            this.entries.forEach(element => {
                if (element.date >= from, element.date <= to) {
                    newEntries.push(element);
                }
            });
            return new Sales(newEntries);
        }

        totalSalesDictionary(): {[id: string] : any} {
            var item =
            {
                'restaurant': this.totalRestaurantSales(),
                'delivery': this.totalDeliverySales(),
                'deliveroo': this.totalDeliverooTurnover(),
                'uber': this.totalUberTurnover(),
                'kitchen': this.totalKitchenSales(),
                'grill': this.totalGrillSales(),
                'service': this.totalServiceSales(),
                'cash': this.totalCashTurnover(),
                'card': this.totalCardTurnover(),
                'seamless': this.totalSeamlessTurnover(),
                'mealpal': this.totalMealPalTurnover(),
                'total tax': this.totalSalesTax(),
                'total': this.totalSales(),
                'food cost' : this.totalFoodCost(),
                'total cost' : this.totalCosts(),

            }
            return item;
        }

        add(entry: Entry) {
            this.entries.push(entry);
        }

        totalAmount() {
            var total = 0.0;
            this.entries.forEach(element => {
                if (element.status.active()) {
                    total += element.amount;
                }
            })
            return total;
        }

        totalSales() {
            var total = 0.0;
            this.entries.forEach(element => {
                if (element.status.active() && element.account.isSales()) {
                    total += element.amount;
                }
            })
            return total;
        }

        totalSalesTax() {
            var total = 0.0;
            this.entries.forEach(element => {
                if (element.status.active() && element.account.isSales()) {
                    total += element.taxAmount;
                }
            })
            return total;
        }

        totalKitchenSales() {
            var total = 0.0;
            this.entries.forEach(element => {
                if (element.status.active() && element.category.isKitchen()) {
                    total += element.amount;
                }
            })
            return total;
        }
        totalGrillSales() {
            var total = 0.0;
            this.entries.forEach(element => {
                if (element.status.active() && element.category.isGrill()) {
                    total += element.amount;
                }
            })
            return total;
        }

        totalServiceSales() {
            var total = 0.0;
            this.entries.forEach(element => {
                if (element.status.active() && element.category.isGrill()) {
                    total += element.amount;
                }
            })
            return total;
        }

        totalRestaurantSales() {
            var total = 0.0;
            this.entries.forEach(element => {
                if (element.status.active() && element.account.isRestaurant()) {
                    total += element.amount;
                }
            })
            return total;
        }
        totalDeliverySales() {
            var total = 0.0;
            this.entries.forEach(element => {
                if (element.status.active() && element.account.isDelivery() && element.isPosSale()) {
                    total += element.amount;
                }
            })
            return total;
        }
        totalDeliverooTurnover() {
            var total = 0.0;
            this.entries.forEach(element => {
                if (element.status.active() && element.account.isDeliveroo() && element.isPosSale()) {
                    total += element.amount;
                }
            })
            return total;
        }
        totalUberTurnover() {
            var total = 0.0;
            this.entries.forEach(element => {
                if (element.status.active() && element.account.isUber() && element.isPosSale()) {
                    total += element.amount;
                }
            })
            return total;
        }

        totalCashTurnover() {
            var total = 0.0;
            this.entries.forEach(element => {
                if (element.status.active() && element.account.isCash() && element.isPosSale()) {
                    total += element.amount;
                }
            })
            return total;
        }
        totalCardTurnover() {
            var total = 0.0;
            this.entries.forEach(element => {
                if (element.status.active() && element.account.isCard() && element.isPosSale()) {
                    total += element.amount;
                }
            })
            return total;
        }

        totalSeamlessTurnover() {
            var total = 0.0;
            this.entries.forEach(element => {
                if (element.status.active() && element.account.isSeamless() && element.isPosSale()) {
                    total += element.amount;
                }
            })
            return total;
        }
        totalMealPalTurnover() {
            var total = 0.0;
            this.entries.forEach(element => {
                if (element.status.active() && element.account.isMealPal() && element.isPosSale()) {
                    total += element.amount;
                }
            })
            return total;
        }

        totalFoodCost() {
            var total = 0.0;
            this.entries.forEach(element => {
                if (element.status.active() && element.account.isFoodCost()) {
                    total += element.amount;
                }
            })
            return total;
        }

        totalCosts() {
            var total = 0.0;
            this.entries.forEach(element => {
                if (element.status.active() && (element.type.isExpense())) {
                    total += element.amount;
                }
            })
            return total;
        }

        public static fromAccountDataTable(data: SheetHeadedData): Sales {
            var entries: Entry[] = [];
            data.getValues();
            data.values.forEach(element => {
                var entry = Entry.fromXeroItem(element);
                if (entry.date) { // Check and remove empty entries
                    entries.push(entry);
                }
            });
            return new Sales(entries);
        }
    }

    export class DaySales {
        days: { [id: string]: Sales };
        constructor(sales: Sales) {
            var dateDictionary: { [id: string]: Sales } = {};
            sales.entries.forEach(element => {
                if (element.date) {
                    var key = dayString(element.date);
                    if (!dateDictionary[key]) {
                        dateDictionary[key] = new Sales([]);
                    }
                    dateDictionary[key].add(element);    
                }
            });
            this.days = dateDictionary;
        }
    }

    function dayString(date: Date): string {
        var timezone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();

        let formatted_date = Utilities.formatDate(date, timezone.toString(), 'yyyy/MM/dd')
        return formatted_date;
    }

    export function dayStringToDate(dayString: string): Date {
        var stringValue = (dayString.replace("/", "-").replace("/", "-")) + "T00:00:00+00:00";
        var date = new Date(Date.parse(stringValue));
        return date;
    }

    // Applies leading 0. Used for date formatting
    function appendLeadingZeroes(n) {
        if (n <= 9) {
            return "0" + n;
        }
        return n
    }



    export class Entry {
        lineItemId: string;
        date: Date;
        contactName: string;
        status: Status;
        amount: number;
        taxAmount: number;
        category: TrackingCategory;
        account: Account;
        type: EntryType;


        constructor(
            lineItemId: string,
            date: Date,
            contactName: string,
            status: Status,
            amount: number,
            taxAmount: number,
            category: TrackingCategory,
            account: Account,
            type: EntryType
            ) {

            this.lineItemId = lineItemId;
            this.date = date;
            this.contactName = contactName;
            this.status = status;
            this.amount = amount;
            this.taxAmount = taxAmount;
            this.category = category;
            this.account = account;
            this.type = type;
        }

        toItem(): { [id: string]: any } {
            var item =
            {
                'id': this.lineItemId,
                'date': this.date,
                'contactName': this.contactName,
                'status': this.status.status,
                'amount': this.amount,
                'tax amount': this.taxAmount,
                'category': this.category.category,
                'type': this.type.type,
                'account': this.account.account,
                'kitchen': this.category.isKitchen(),
                'sales': this.account.isSales(),
                'delivery': this.account.isDelivery(),
                'restaurant': this.account.isRestaurant()
            }
            return item;
        }

        isPosSale() : Boolean {
            return (this.contactName === 'iKentoo' || this.contactName === 'POS Sales' || this.contactName === 'Delivery Sales');
        }

        isFoodCost() : Boolean {
            return this.account.isFoodCost();
        }

        isExpense(): Boolean {
            if (this.isFoodCost) {
                return false;
            }
            return this.type.isExpense();
        }

        public static fromXeroItem(item: { [id: string]: any }): Entry {
            var id: string = item['Line Item ID'];
            var statusString: string = item['Status'];
            var status = new Status(statusString);
            var accountString: string = item['Account'];
            var descriptionString: string = item['Description'];
            var contactName: string = item['Contact Name'];
            var account = new Account(accountString, descriptionString);
            var date: Date = item['Date'];
            var amount: number = item['GBP Amount - No Tax'];
            var taxAmount: number = item['Tax Amount'];  // TODO: Should be changed to GBP Tax amount when ready
            var categoryString: string = item['Tracking Category'];
            var category = new TrackingCategory(categoryString);
            var typeString: string = item['Type'];
            var type = new EntryType(typeString);
            return new Entry(id, date, contactName, status, amount, taxAmount, category, account, type);
        }
    }
}

function updateDayTotalsTable() {
    SpreadsheetApp.getActive().toast("1. Getting Accounting Data");
    var dataTable = new SheetHeadedData(AllAccountingDataTable, new SSHeadedRange(0,0,0,0,9,10));
    SpreadsheetApp.getActive().toast("2. Processing Accounting Data");
    var sales = Accounting.Sales.fromAccountDataTable(dataTable);
    SpreadsheetApp.getActive().toast("3. Dividing to Dates");
    var dailySales = new Accounting.DaySales(sales);
    var exportTable = new SheetHeadedData(SalesDataTable, new SSHeadedRange(0,0,0,0,1,2));
    SpreadsheetApp.getActive().toast("4. Clearing Table");
    exportTable.clearValues();
    var sortedDays = Object.keys(dailySales.days).sort().reverse();
    var i = 0;
    SpreadsheetApp.getActive().toast("5. Preparing Results "+i+"/"+sortedDays.length);
    var data: [{ [id: string]: any }] = [{}];
    data.pop();

    sortedDays.forEach(key => {
        var daySales = dailySales.days[key];
        var item = daySales.totalSalesDictionary();
        item["id"] =  key;
        item["date"] = Accounting.dayStringToDate(key);
        data.push(item);
        i++;
        if (i%15 == 0) {
            SpreadsheetApp.getActive().toast("5. Preparing Results "+i+"/"+sortedDays.length);
        }
    })
    SpreadsheetApp.getActive().toast("6. Displaying Data");
    exportTable.values = data;
    exportTable.writeValues();
    SpreadsheetApp.getActive().toast("7. Complete");
}

function updateDayTotals(from: Date, to: Date, headerString: string): any[][]{
    var headers = headerString.toString().split(',');
    var data = [];

    var dataTable = new SheetHeadedData(AllAccountingDataTable, new SSHeadedRange(0,0,0,0,9,10));
    var sales = Accounting.Sales.fromAccountDataTable(dataTable);
    sales = sales.salesForPeriod(from, to);
    var dailySales = new Accounting.DaySales(sales);
    var sortedDays = Object.keys(dailySales.days).sort().reverse();
    sortedDays.forEach(key => {
        var daySales = dailySales.days[key];
        var item = daySales.totalSalesDictionary();
        item["id"] =  key;
        item["date"] = Accounting.dayStringToDate(key);
        var values = [];
        headers.forEach(key => {
            values.push(item[key] || "");
        });
        data.push(values);
    })
    return data;
}

function sameDay(day1: Date, day2: Date): Boolean {
    if (day1.getFullYear() != day2.getFullYear()) {
        return false;
    }
    if (day1.getMonth() != day2.getMonth()) {
        return false;
    }
    if (day1.getDay() != day2.getDay()) {
        return false;
    }
    return true;
}
function containsSameDay(day: Date, dayArray: Date[]): Boolean {
    var found = false;
    dayArray.forEach(element => {
        if (sameDay(element, day)) {
            found = true;
            return;
        }
    });
    return found;
}
