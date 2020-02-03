class Func {
	public capitals = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
	public smalls = "abcdefghijklmnopqrstuvwxyz";
	public digits = "1234567890";
	public symbols = ",./?'!@#$%^&*()-_+=`~\\| ";
	public months = ['January', 'Febuary', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
	public days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
	public genders = ['Male', 'Female', 'Do not disclose'];
	public maritals = ['Married', 'Single', 'Divorced'];
	public religions = ['Christainity', 'Islam', 'Judaism', 'Paganism', 'Budism'];
	public userTypes = ['student', 'staff', 'admin', 'ceo'];
	public staffRequests = ['leave', 'allowance'];
	public studentsRequests = ['absence', 'academic'];
	public subjectList = ['Mathematics', 'English', 'Physics', 'Chemistry', 'Biology', 'Agriculture', 'Literature', 'History'].sort();
	public subjectLevels = ['General', 'Senior', 'Science', 'Arts', 'Junior'];
	public fontStyles = ['Arial', 'Times New Roman', 'Helvetica', 'Times', 'Courier New', 'Verdana', 'Courier', 'Arial Narrow', 'Candara', 'Geneva', 'Calibri', 'Optima', 'Cambria', 'Garamond', 'Perpetua', 'Monaco', 'Didot', 'Brush Script MT', 'Lucida Bright', 'Copperplate', 'Serif', 'San-Serif', 'Georgia', 'Segoe UI'];
	public pixelSizes = ['1px', '2px', '3px', '4px', '5px', '6px', '7px', '8px', '9px', '10px', '20px', '30px', '40px', '50px', '60px', '70px', '80px', '90px', '100px', 'None', 'Unset'];
	public colors = ['Red', 'Green', 'Blue', 'Yellow', 'Black', 'White', 'Purple', 'Violet', 'Indigo', 'Orange', 'Transparent', 'None', 'Unset'];
	public boldness = [100, 200, 300, 400, 500, 600, 700, 800, 900, 1000, 'lighter', 'bold', 'bolder', 'normal', 'unset'];
	public borderTypes = ['Solid', 'Dotted', 'Double', 'Groove', 'Dashed', 'Inset', 'None', 'Unset', 'Outset', 'Rigged', 'Inherit', 'Initial'];
	public shadows = ['2px 2px 5px 2px red', '2px 2px 5px green', '2px 2px yellow', '2px black', 'None', 'Unset'];
	public borders = ['1px solid black', '2px dotted green', '3px dashed yellow', '1px double red', 'None', 'Unset'];

	constructor() {

	}

	public extractFromJsonArray(meta, source) {
		let keys: any = Object.keys(meta);
		// @ts-ignore
		let values: any = Object.values(meta);
		let eSource = [];
		if (this.isset(source)) {
			for (let obj of source) {
				let object = {};
				eSource.push(object);

				for (let i in keys) {
					// @ts-ignore					
					if (Object.keys(obj).includes(values[i])) {
						object[keys[i]] = obj[values[i]];
					}
				}
			}
			return eSource;
		}
	}

	public trimMonthArray() {
		let trimmedMonths = [];
		for (let i = 0; i < this.months[i].length; i++) {
			trimmedMonths.push(this.months[i].slice(0, 3));
		}
		return trimmedMonths;
	}

	public jsStyleName(name) {
		let newName = '';
		for (let i = 0; i < name.length; i++) {
			if (name[i] == '-') {
				i++;
				newName += name[i].toUpperCase();
			}
			else newName += name[i].toLowerCase();
		}
		return newName;
	}

	public cssStyleName(name) {
		let newName = '';
		for (let i = 0; i < name.length; i++) {
			if (this.isCapital(name[i])) newName += '-';
			newName += name[i].toLowerCase();
		}

		return newName;
	}

	public edittedUrl(params) {
		var url = this.urlSplitter(params.url);
		url.vars[params.toAdd] = params.addValue.toLowerCase();
		return this.urlMerger(url, params.toAdd);
	}

	public hasArrayElement(haystack, needle) {
		for (var i of needle) {
			if (haystack.indexOf(i) != -1) return true;
		}
		return false;
	}

	public removeDuplicate(haystack) {
		var single = [];
		for (var x in haystack) {
			if (!this.hasString(single, haystack[x])) {
				single.push(haystack[x]);
			}
		}
		return single;
	}

	public addCommaToMoney(money) {
		let inverse = '',
			position;
		for (let i = money.length - 1; i >= 0; i--) {
			inverse += money[i];
		}

		money = "";

		for (let i = 0; i < inverse.length; i++) {
			position = (i + 1) % 3;
			money += inverse[i];
			if (position == 0) {
				if (i != inverse.length - 1) {
					money += ',';
				}
			}
		}
		inverse = '';

		for (let i = money.length - 1; i >= 0; i--) {
			inverse += money[i];
		}
		return inverse;
	}

	public isCapital(value) {
		if (value.length == 1) {
			return this.isSubString(this.capitals, value);
		}
	}

	public capitalize(value) {
		if (!this.isCapital(value[0])) {
			value = value.split('');
			value[0] = this.capitals[this.smalls.indexOf(value[0])];
			return this.stringReplace(value.toString(), ',', '');
		}
		return value;
	}

	public isSmall(value) {
		if (value.length == 1) {
			return this.isSubString(this.smalls, value);
		}
	}

	public isSymbol(value) {
		if (value.length == 1) {
			return this.isSubString(this.symbols, value);
		}
	}

	public isName(value) {
		for (var x in value) {
			if (this.isDigit(value[x])) {
				return false;
			}
		}
		return true;
	}

	public isNumber(value) {
		for (var x in value) {
			if (!this.isDigit(value[x]) && value[x] != '.') {
				return false;
			}
		}
		return value;
	}

	public isPasswordValid(value) {
		var len = value.length;
		if (len > 7) {
			for (var a in value) {
				if (this.isCapital(value[a])) {
					for (var b in value) {
						if (this.isSmall(value[b])) {
							for (var c in value) {
								if (this.isDigit(value[c])) {
									for (var d in value) {
										if (this.isSymbol(value[d])) {
											return true;
										}
									}
								}
							}
						}
					}
				}
			}
		}
		return false;
	}

	public isSubString(haystack, value) {
		if (haystack.indexOf(value) != -1) return true;
		return false;
	}

	public isTimeValid(time) {
		time = time.split(':');

		if (time.length == 2 || time.length == 3) {
			let hour: number = time[0];
			let minutes: number = time[1];
			let seconds: number = time[2] || 0;
			let total: number = 0;

			if (time.length == 3) {
				if (hour > 23 || hour < 0 || minutes > 59 || minutes < 0 || seconds > 59 || seconds < 0) {
					return false;
				}
			} else {
				if (hour > 23 || hour < 0 || minutes > 59 || minutes < 0) {
					return false;
				}
			}

			hour = this.secondsInHours(hour);

			minutes = this.secondsInMinutes(minutes);

			total = hour + minutes + Math.floor(seconds);

			return total;
		}
		return false;
	}

	public isDigit(value) {
		value = new String(value);
		if (value.length == 1) {
			return this.isSubString(this.digits, value);
		}
		return false;
	}

	public isEmail(value) {
		var email_parts = value.split('@');
		if (email_parts.length != 2) {
			return false;
		} else {
			if (this.isSpaceString(email_parts[0])) {
				return false;
			}
			var dot_parts = email_parts[1].split('.');
			if (dot_parts.length != 2) {
				return false;
			} else {
				if (this.isSpaceString(dot_parts[0])) {
					return false;
				}
				if (this.isSpaceString(dot_parts[1])) {
					return false;
				}
			}
		}
		return true;
	}

	public isDateValid(value) {
		if (this.isDate(value)) {
			if (this.isYearValid(value)) {
				if (this.isMonthValid(value)) {
					if (this.isDayValid(value)) {
						return true;
					}
				}
			}
		}
		return false;
	}

	public isDayValid(value) {
		var v_day: any = '';
		for (var i = 0; i < 2; i++) {
			if (this.isset(value[i + 8])) v_day += value[i + 8];
		}

		var limit = 0;
		var month: any = this.isMonthValid(value);

		if (month == '01') {
			limit = 31;
		} else if (month == '02') {
			if (this.isLeapYear(this.isYearValid(value))) {
				limit = 29;
			} else {
				limit = 28;
			}
		} else if (month == '03') {
			limit = 31;
		} else if (month == '04') {
			limit = 30;
		} else if (month == '05') {
			limit = 31;
		} else if (month == '06') {
			limit = 30;
		} else if (month == '07') {
			limit = 31;
		} else if (month == '08') {
			limit = 31;
		} else if (month == '09') {
			limit = 30;
		} else if (month == '10') {
			limit = 31;
		} else if (month == '11') {
			limit = 30;
		} else if (month == '12') {
			limit = 31;
		}

		if (limit < v_day) {
			return 0;
		}
		return v_day;
	}

	public generateRandom(length) {
		var string = this.capitals + this.smalls + this.digits;
		var alphanumeric = '';
		for (var i = 0; i < length; i++) {
			alphanumeric += string[Math.floor(Math.random() * string.length)];
		}
		return alphanumeric;
	}

	public isDate(value) {

		var len = value.length;
		if (len == 10) {
			for (let x = 0; x < len; x++) {
				if (this.isDigit(value[x])) {
					continue;
				} else {
					if (x === 4 || x == 7) {
						if (value[x] == '-') {
							continue;
						} else {
							return false;
						}
					} else {
						return false;
					}
				}
			}
		} else {
			return false;
		}
		return true;
	}

	public isMonthValid(value) {
		var v_month: any = '';
		for (var i = 0; i < 2; i++) {
			if (this.isNumber(value[i + 5])) v_month += value[i + 5];
		}
		if (v_month > 12 || v_month < 1) {
			return 0;
		}
		return v_month;
	}

	public isYearValid(value) {
		var year = new Date().getFullYear();
		var v_year: number;
		for (var i = 0; i < 4; i++) {
			if (this.isNumber(value[i + 0])) v_year += value[i + 0];
		}
		if (v_year > year) {
			return 0;
		}
		return v_year;
	}

	public getYear(value) {
		var v_year: any = '';

		for (var i = 0; i < 4; i++) {
			v_year += value[i];
		}
		return v_year;
	}

	public isLeapYear(value) {
		if (value % 4 == 0) {
			if ((value % 100 == 0) && (value % 400 != 0)) {
				return false;
			}
			return true;
		}
		return false;
	}

	public daysInMonth(month, year) {
		var days = 0;
		if (month == '01') {
			days = 31;
		} else if (month == '02') {
			if (this.isLeapYear(year)) {
				days = 29;
			} else {
				days = 28;
			}
		} else if (month == '03') {
			days = 31;
		} else if (month == '04') {
			days = 30;
		} else if (month == '05') {
			days = 31;
		} else if (month == '06') {
			days = 30;
		} else if (month == '07') {
			days = 31;
		} else if (month == '08') {
			days = 31;
		} else if (month == '09') {
			days = 30;
		} else if (month == '10') {
			days = 31;
		} else if (month == '11') {
			days = 30;
		} else if (month == '12') {
			days = 31;
		}

		return days;
	}

	public dateValue(date) {
		var value: number = 0;

		var year: number = this.getYear(date) * 365;

		var month: number = 0;
		for (var i = 1; i < this.isMonthValid(date); i++) {
			month = this.daysInMonth(i, this.getYear(date));
		}
		var day: number = this.isDayValid(date);

		value = year + month + day;

		return value;
	}

	public getDateObject(value: any) {
		let days: number = Math.floor(value / this.secondsInDays(1));
		// console.log(value, this.secondsInDays(1));

		value -= this.secondsInDays(days);

		let hours = Math.floor(value / this.secondsInHours(1));
		value -= this.secondsInHours(hours);

		let minutes = Math.floor(value / this.secondsInMinutes(1));
		value -= this.secondsInMinutes(minutes);

		let seconds = value;

		return { days, hours, minutes, seconds };
	}

	public today() {
		var today: any = new Date;
		var month: any = (today.getMonth() / 1 + 1).toString();
		if (month.length != 2) {
			month = '0' + month;
		}
		today = (today.getFullYear()) + '-' + month + '-' + today.getDate();
		return today;
	}

	public timeToday() {
		let date = new Date();
		let hour: any = date.getHours();
		let minutes: any = date.getMinutes();
		let seconds: any = date.getSeconds();

		let time = this.isTimeValid(`${hour}:${minutes}:${seconds}`);
		return time ? time : -1;
	}

	public time() {
		let date = new Date();
		let hour: any = date.getHours();
		let minutes: any = date.getMinutes();
		let seconds: any = date.getSeconds();

		return `${hour}:${minutes}:${seconds}`;
	}

	public objectLength(object) {
		return Object.keys(object).length;
	}

	public getObjectArrayKeys(array) {
		let keys: any = [];
		for (let object of array) {
			for (let i of Object.keys(object)) {
				if (!keys.includes(i)) {
					keys.push(i);
				}
			}
		}

		return keys;
	}

	public dateWithToday(date) {
		var today = Math.floor(this.dateValue(this.today()));
		let dateValue = Math.floor(this.dateValue(date));

		var value = { diff: (dateValue - today), when: '' };
		if (dateValue > today) {
			value.when = 'future';
		}
		else if (dateValue == today) {
			value.when = 'today';
		}
		else {
			value.when = 'past';
		}
		return value;
	}

	public dateString(date) {
		var year = new Number(this.getYear(date));
		var month: any = new Number(this.isMonthValid(date));
		var day = new Number(this.isDayValid(date));

		return day + ' ' + this.months[month - 1] + ', ' + year;
	}

	public copyFormData(formData) {
		let myFormData: any = {};
		try {
			for (let [key, value] of formData.entries()) {
				myFormData[key] = value;
			}
		} catch (error) {
			console.log(formData.serializeArray());
			return null;
		}
		return myFormData;
	}

	public isSpaceString(value) {
		if (value == '') {
			return true;
		} else {
			for (var x in value) {
				if (value[x] != ' ') {
					return false;
				}
			}
		}
		return true;
	}

	public getOccurancesOf(haystack, needle) {
		let found = [];
		for (let key = 0; key < haystack.length; key++) {
			if (haystack[key] == needle) found.push(key);
		}
		return found;
	}

	public deleteOccuranceOf(haystack, needle) {
		var isArray = Array.isArray(haystack);
		var value: any = (isArray) ? [] : '';
		for (var i of haystack) {
			if (i == needle) continue;
			(isArray) ? value.push(i) : value += i;
		}
		return value;
	}

	public deleteArrayInPosition(haystack, position) {
		var tmp = [];
		for (var i = 0; i < haystack.length; i++) {
			if (i == position) {
				continue;
			}
			tmp.push(haystack[i]);
		}
		return tmp;
	}

	public insertArrayInPosition(haystack, needle, insert) {
		var position: any = this.getPositionOfArray(haystack, needle);
		var tmp = [];
		for (var i = 0; i < haystack.length; i++) {
			tmp.push(haystack[i]);
			if (i === position) {
				tmp.push(insert);
			}
		}
		return tmp;
	}

	public getPositionOfArray(haystack, needle) {
		for (var x in haystack) {
			if (JSON.stringify(haystack[x]) == JSON.stringify(needle)) {
				return x;
			}
		}
		return false;
	}

	public getPositionInArray(haystack, needle) {
		for (var x in haystack) {
			if (haystack[x] == needle) {
				return x;
			}
		}
		return -1;
	}

	public hasArray(haystack, needle) {
		haystack = JSON.stringify(haystack);
		needle = JSON.stringify(needle);

		return (haystack.indexOf(needle) >= 0) ? true : false;
	}

	public hasString(haystack, needle) {
		for (var x in haystack) {
			if (needle == haystack[x]) {
				return true;
			}
		}
		return false;
	}

	public trem(needle) {
		//remove the prepended spaces
		if (needle[0] == ' ') {
			let new_needle = '';
			for (let i = 0; i < needle.length; i++) {
				if (i != 0) {
					new_needle += needle[i];
				}
			}
			needle = this.trem(new_needle);
		}

		//remove the appended spaces
		if (needle[needle.length - 1] == ' ') {
			let new_needle = '';
			for (let i = 0; i < needle.length; i++) {
				if (i != needle.length - 1) {
					new_needle += needle[i];
				}
			}
			needle = this.trem(new_needle);
		}
		return needle;
	}

	public stringReplace(word, from, to) {
		var returnvalue = '';
		for (var x in word) {
			if (word[x] == from) {
				returnvalue += to;
				continue;
			}
			returnvalue += word[x];
		}
		return returnvalue;
	}

	public converToRealPath(path) {
		if (path[path.length - 1] != '/') {
			path += '/';
		}
		return path;
	}

	public isSpacialCharacter(char) {
		var specialcharacters = "'\\/:?*<>|!.";
		for (var i = 0; i < specialcharacters.length; i++) {
			if (specialcharacters[i] == char) {
				return true;
			}
		}
		return false;
	}

	public countChar(haystack, needle) {
		var j = 0;
		for (var i = 0; i < haystack.length; i++) {
			if (haystack[i] == needle) {
				j++;
			}
		}
		return j;
	}

	public isset(variable) {
		let result = false;
		try {
			result = (typeof variable !== 'undefined');
		} catch (error) {
			console.log(error);
		}
		return result;
	}

	public isnull(variable) {
		return variable == null;
	}

	public setNotNull(variable) {
		return this.isset(variable) && !this.isnull(variable);
	}

	public urlMerger(splitUrl, lastQuery) {
		var hostType = (this.isset(splitUrl.hostType)) ? splitUrl.hostType : 'http';
		var hostName = (this.isset(splitUrl.hostName)) ? splitUrl.hostName : '';
		var port = (this.isset(splitUrl.host)) ? splitUrl.port : '';
		var pathName = (this.isset(splitUrl.pathName)) ? splitUrl.pathName : '';
		var queries = '?';
		var keepMapping = true;
		if (this.isset(splitUrl.vars)) {
			Object.keys(splitUrl.vars).map(key => {
				if (keepMapping) queries += key + '=' + splitUrl.vars[key] + '&';
				if (key == lastQuery) keepMapping = false;
			});
		}
		var location = hostType + '::/' + hostName + ':' + port + '/' + pathName + queries;
		location = (location.lastIndexOf('&') == location.length - 1) ? location.slice(0, location.length - 1) : location;
		location = (location.lastIndexOf('=') == location.length - 1) ? location.slice(0, location.length - 1) : location;
		return location;
	}

	public urlSplitter(location) {
		if (this.isset(location)) {
			location = location.toString();
			var httpType = (location.indexOf('://') === -1) ? null : location.split('://')[0];
			var fullPath = location.split('://').pop(0);
			var host = fullPath.split('/')[0];
			var hostName = host.split(':')[0];
			var port = host.split(':').pop(0);
			var path = '/' + fullPath.split('/').pop(0);
			var pathName = path.split('?')[0];
			var queries = (path.indexOf('?') === -1) ? null : path.split('?').pop();

			var vars = {};
			if (queries != null) {
				var query = queries.split('&');
				for (var x in query) {
					var parts = query[x].split('=');
					if (parts[1]) {
						vars[this.stringReplace(parts[0], '-', ' ')] = this.stringReplace(parts[1], '-', ' ');
					} else {
						vars[this.stringReplace(parts[0], '-', ' ')] = '';
					}
				}
			}
			var httphost = httpType + '://' + host;
			return { location: location, httpType: httpType, fullPath: fullPath, host: host, httphost: httphost, hostName: hostName, port: port, path: path, pathName: pathName, queries: queries, vars: vars };
		}
	}

	public getUrlVars(location) {
		location = location.toString();
		var queries = (location.indexOf('?') === -1) ? null : location.split('?').pop(0);
		var vars = {};

		if (queries != null) {
			var query = queries.split('&');
			for (var x in query) {
				var parts = query[x].split('=');
				if (parts[1]) {
					vars[this.stringReplace(parts[0], '-', ' ')] = this.stringReplace(parts[1], '-', ' ');
				} else {
					vars[this.stringReplace(parts[0], '-', ' ')] = '';
				}
			}
		}
		return vars;
	}

	public objectToArray(obj) {
		var arr = [];
		Object.keys(obj).map((key) => {
			arr[key] = obj[key];
		});
		return arr;
	}

	public async runParallel(functions, callBack) {
		var results = {};
		for (var f in functions) {
			results[f] = await functions[f];
		}
		callBack(results);
	}

	public isMobile() {
		return (/Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent));
	}

	public secondsInDays(days) {
		return Math.floor(days * 24 * 60 * 60);
	}

	public secondsInHours(hours) {
		return Math.floor(hours * 60 * 60);
	}

	public secondsInMinutes(minutes) {
		return Math.floor(minutes * 60);
	}
}

let func = new Func();

export default func;