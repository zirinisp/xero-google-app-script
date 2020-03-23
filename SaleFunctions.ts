
import "./table.js";
import "./item.js";
import "./selector.js";



var AllAccountingDataTable = 'All Accounting Data';
var SalesDataTable = 'Day Sales';//'All Accounting Data';


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
        Other = ''
    }

    class Account {
        account: string;
        constructor(account: string) {
            this.account = account;
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
            if (this.account.toString() === AccountEnum.UberEats) {
                return true;
            }
            return false;
        }

        isDeliveroo(): Boolean {
            if (this.account.toString() === AccountEnum.Deliveroo) {
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
                'total': this.totalSales()
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
                if (element.status.active() && element.account.isDelivery()) {
                    total += element.amount;
                }
            })
            return total;
        }
        totalDeliverooTurnover() {
            var total = 0.0;
            this.entries.forEach(element => {
                if (element.status.active() && element.account.isDeliveroo()) {
                    total += element.amount;
                }
            })
            return total;
        }
        totalUberTurnover() {
            var total = 0.0;
            this.entries.forEach(element => {
                if (element.status.active() && element.account.isUber()) {
                    total += element.amount;
                }
            })
            return total;
        }

        public static fromAccountDataTable(table: Table): Sales {
            var entries: Entry[] = [];
            table.items.forEach(element => {
                var entry = Entry.fromItem(element);
                entries.push(entry);
            });
            return new Sales(entries);
        }
    }

    export class DaySales {
        days: { [id: string]: Sales };
        constructor(sales: Sales) {
            var dateDictionary: { [id: string]: Sales } = {};
            sales.entries.forEach(element => {
                var key = dayString(element.date);
                if (!dateDictionary[key]) {
                    dateDictionary[key] = new Sales([]);
                }
                dateDictionary[key].add(element);
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
        status: Status;
        amount: number;
        category: TrackingCategory;
        account: Account;


        constructor(
            lineItemId: string,
            date: Date,
            status: Status,
            amount: number,
            category: TrackingCategory,
            account: Account) {

            this.lineItemId = lineItemId;
            this.date = date;
            this.status = status;
            this.amount = amount;
            this.category = category;
            this.account = account;
        }

        toItem(): { [id: string]: any } {
            var item =
            {
                'id': this.lineItemId,
                'date': this.date,
                'status': this.status.status,
                'amount': this.amount,
                'category': this.category.category,
                'account': this.account.account,
                'kitchen': this.category.isKitchen(),
                'sales': this.account.isSales(),
                'delivery': this.account.isDelivery(),
                'restaurant': this.account.isRestaurant()
            }
            return item;
        }

        public static fromItem(item: Item): Entry {
            var id: string = item.getFieldValue('Line Item ID');
            var statusString: string = item.getFieldValue('Status');
            var status = new Status(statusString);
            var accountString: string = item.getFieldValue('Account');
            var account = new Account(accountString);
            var date: Date = item.getFieldValue('Date');
            var amount: number = item.getFieldValue('GBP Amount - No Tax');
            var categoryString: string = item.getFieldValue('Tracking Category');
            var category = new TrackingCategory(categoryString);
            return new Entry(id, date, status, amount, category, account);
        }
    }
}

function updateSales() {
    var dataTable = getTable(AllAccountingDataTable, 9, 'Line Item ID');
    var sales = Accounting.Sales.fromAccountDataTable(dataTable);
    var dailySales = new Accounting.DaySales(sales);
    var exportTable = getTable(SalesDataTable, 1, 'id');
    exportTable.deleteAll();
    var sortedDays = Object.keys(dailySales.days).sort().reverse();
    sortedDays.forEach(key => {
        var daySales = dailySales.days[key];
        daySales.entries.forEach(element => {
            exportTable.add(element.toItem());
        })
    });
    exportTable.commit();
}

function updateDayTotalsTable() {
    SpreadsheetApp.getActive().toast("1. Getting Accounting Data");
    var dataTable = getTable(AllAccountingDataTable, 9, 'Line Item ID');
    SpreadsheetApp.getActive().toast("2. Processing Accounting Data");
    var sales = Accounting.Sales.fromAccountDataTable(dataTable);
    SpreadsheetApp.getActive().toast("3. Dividing to Dates");
    var dailySales = new Accounting.DaySales(sales);
    var exportTable = getTable('Test Export 2', 1, 'id');
    SpreadsheetApp.getActive().toast("4. Clearing Table");
    exportTable.deleteAll();
    var sortedDays = Object.keys(dailySales.days).sort().reverse();
    var i = 0;
    SpreadsheetApp.getActive().toast("5. Preparing Results "+i+"/"+sortedDays.length);
    sortedDays.forEach(key => {
        var daySales = dailySales.days[key];
        var item = daySales.totalSalesDictionary();
        item["id"] =  key;
        item["date"] = Accounting.dayStringToDate(key);
        exportTable.add(item);
        i++;
        if (i%15 == 0) {
            SpreadsheetApp.getActive().toast("5. Preparing Results "+i+"/"+sortedDays.length);
        }
    })
    SpreadsheetApp.getActive().toast("6. Displaying Data");
    exportTable.commit();
    SpreadsheetApp.getActive().toast("7. Complete");
}

function updateDayTotals(from: Date, to: Date, headerString: string): any[][]{
    var headers = headerString.toString().split(',');
    var data = [];

    var dataTable = getTable(AllAccountingDataTable, 9, 'Line Item ID');
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
