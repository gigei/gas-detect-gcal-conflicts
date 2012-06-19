
// Copyright (C) 2012 Nihon Gigei, Inc. <http://rakumo.gigei.jp/>
// Licensed under the terms of the MIT License

var Config = function() {
	this.DOMAIN = ScriptProperties.getProperty('DOMAIN');
	this.MAX_DATE = ScriptProperties.getProperty('MAX_DATE');
	this.ALERT_MAIL_ADDRESS = ScriptProperties.getProperty('ALERT_MAIL_ADDRESS');
	var strIsSend = ScriptProperties.getProperty('IS_SEND_MAIL_CREATOR'),
		isSend = undefined;
	if (strIsSend == 'true') {
		isSend = true;
	} else if (strIsSend == 'false') {
		isSend = false;
	}
	this.IS_SEND_MAIL_CREATOR = isSend;
};
Config.prototype = {
	DOMAIN: undefined,
	MAX_DATE: undefined,
	ALERT_MAIL_ADDRESS: undefined,
	IS_SEND_MAIL_CREATOR: undefined
};

var GoogleOAuth = function(name, scope) {
	this.name = name;
	this.scope = scope;
};
GoogleOAuth.prototype = {
	name: undefined,

	scope: undefined,

	initialize: function() {
		var oAuthConfig = UrlFetchApp.addOAuthService(this.name);
		oAuthConfig.setRequestTokenUrl('https://www.google.com/accounts/OAuthGetRequestToken?scope=' + this.scope);
		oAuthConfig.setAuthorizationUrl('https://www.google.com/accounts/OAuthAuthorizeToken');
		oAuthConfig.setAccessTokenUrl('https://www.google.com/accounts/OAuthGetAccessToken');
		oAuthConfig.setConsumerKey('anonymous');
		oAuthConfig.setConsumerSecret('anonymous');
	}
};

var SettingUiApp = function() {
	this.config = new Config();
};
SettingUiApp.prototype = {
	config: undefined,

	height: 190,

	width: 285,

	configuration: function() {
		var app = UiApp.createApplication().setHeight(this.height).setWidth(this.width),
			panel = app.createVerticalPanel(),
			grid = app.createGrid(4, 2),
			sheets = SpreadsheetApp.getActiveSpreadsheet();

		app.add(panel);

		this.setLabel(app, panel);
		this.setGrid(app, panel, grid);
		this.setButton(app, panel, grid);

		if (!this.config.DOMAIN || !this.config.MAX_DATE || !this.config.ALERT_MAIL_ADDRESS) {
			app.getElementById('label').setText('未設定');
		}

		sheets.show(app);
	},

	setLabel: function(app, panel) {
		var label = app.createLabel('設定');

		label.setStyleAttribute('font-size', '16pt');
		label.setStyleAttribute('background-color', '#CCCCFF');
		label.setStyleAttribute('text-align', 'center');

		panel.add(label);
	},

	setGrid: function(app, panel, grid) {
		var userAddress = Session.getUser().getEmail(),
			domain = userAddress.split('@');

		grid.setWidget(0, 0, app.createLabel('ドメイン名：'));
		grid.setWidget(0, 1, app.createLabel(domain[1]));

		var textBoxMaxDate = app.createTextBox().setName('max_date');
		textBoxMaxDate.setValue((this.config.MAX_DATE) ?this.config.MAX_DATE :'7');
		grid.setWidget(1, 0, app.createLabel('重複範囲：'));
		grid.setWidget(1, 1, textBoxMaxDate);

		var textBoxAddress = app.createTextBox().setName('address');
		textBoxAddress.setValue((this.config.ALERT_MAIL_ADDRESS) ?this.config.ALERT_MAIL_ADDRESS :userAddress)
		grid.setWidget(2, 0, app.createLabel('通知先メールアドレス：'));
		grid.setWidget(2, 1, textBoxAddress);

		var checkBoxIsSend = app.createCheckBox().setName('is_send');
		checkBoxIsSend.setValue((this.config.IS_SEND_MAIL_CREATOR) ?this.config.IS_SEND_MAIL_CREATOR :false);
		grid.setWidget(3, 0, app.createLabel('予定登録者にも通知：'));
		grid.setWidget(3, 1, checkBoxIsSend);

		panel.add(grid);
	},

	setButton: function(app, panel, grid) {
		var buttonGrid = app.createGrid(1, 2),
			btn,
			label,
			clickHandler;

		btn = app.createButton('設定'),
		clickHandler = app.createServerClickHandler('onClick');
		clickHandler.addCallbackElement(grid);
		btn.addClickHandler(clickHandler);
		buttonGrid.setWidget(0, 0, btn);

		label = app.createLabel().setId('label').setText('').setStyleAttribute("color", "red");
		buttonGrid.setWidget(0, 1, label);

		panel.add(buttonGrid);
	}
};

var CalendarResource = function() {
	this.config = new Config();
	this.articles = ScriptProperties.getProperty('resourcesList_articles');
	this.resources = [];
	this.resourcesLists = [];
	this.get();
};
CalendarResource.prototype = {
	config: undefined,

	aiticles: null,

	resources: [],

	resourcesLists: [],

	limit: 50,

	oAuthName: 'calendarResource',

	oAuthScope: 'https://apps-apis.google.com/a/feeds/calendar/resource/',

	get: function() {
		var scriptStTime = new Date();
		Logger.log('ScriptStTime: ' + scriptStTime.toLocaleString());

		if (!this.config.DOMAIN || !this.config.MAX_DATE || !this.config.ALERT_MAIL_ADDRESS) {
			Logger.log('No Setting Config');
			return;
		}

		if (this.articles) this.resetProperty(this.articles);

		//OAuth
		var oAuth = new GoogleOAuth(this.oAuthName, this.oAuthScope);
		oAuth.initialize();
		//Fetch
		var url = this.oAuthScope + '2.0/' + this.config.DOMAIN + '/',
			fetchArgs = this.getFetchArgs('get', this.oAuthName),
			isNext = true;
		while (isNext == true) {
			Logger.log(url);
			var result = UrlFetchApp.fetch(url, fetchArgs).getContentText(),
				xml = Xml.parse(result);
				next = xml.feed.link[0];

			this.setResources(xml);

			if (next.rel == 'next') {
				url = next.href;
			} else {
				isNext = false;
			}
		}
		this.divideResources();

		var scriptEnTime = new Date();
		Logger.log('ScriptEnTime: ' + scriptEnTime.toLocaleString());
	},

	getFetchArgs: function(method, name) {
		return {
			method: method,
			oAuthServiceName:name,
			oAuthUseToken:'always'
		};
	},

	setResources: function(xml) {
		var entry = xml.feed.entry;
		for (var i = 0; i < entry.length; i++) {
			var resource = entry[i].property;
			this.resources.push([resource[1].value, resource[2].value]);
		}
	},

	divideResources: function() {
		var j = 0;
		this.resourcesLists[j] = [];
		for (var i = 0; i < this.resources.length; i++) {
			var mod = (i != 0) ?i % this.limit :undefined;
			if (mod == 0) {
				j++;
				this.resourcesLists[j] = [];
			}
			this.resourcesLists[j].push(this.resources[i]);
		}
	},

	setProperty: function() {
		this.resetProperty(this.articles);
		for (var i = 0; i < this.resourcesLists.length; i++) {
			var resourcesList = this.resourcesLists[i],
				resourceIds = [],
				strResourceIds = '';
			for (var j = 0; j < resourcesList.length; j++) {
				resourceIds.push(resourcesList[j][1]);
			}
			strResourceIds = resourceIds.join(',');
			ScriptProperties.setProperty('resourcesList_' + String(i), strResourceIds);
		}
		ScriptProperties.setProperty('resourcesList_articles', String(this.resourcesLists.length));
	},

	resetProperty: function(articles) {
		ScriptProperties.setProperty('execute_place', '0');
		for (var i = 0; i < articles; i++) {
			ScriptProperties.setProperty('resourcesList_' + String(i), '');
		}
	}
};

var ClassCalendar = function(email) {
	this.config = new Config();
	this.email = email;
	this.name = null,
	this.events = [];
};
ClassCalendar.prototype = {
	config: undefined,

	email: null,

	name: null,

	events: [],

	oAuthName: 'calendar',

	oAuthScope: 'https://www.google.com/calendar/feeds/',

	fetch: function(min, max) {
		var scriptStTime = new Date();
		Logger.log(scriptStTime.toLocaleString());

		if (!this.config.DOMAIN || !this.config.MAX_DATE || !this.config.ALERT_MAIL_ADDRESS) {
			Logger.log('No Setting Config');
			return;
		}

		//OAuth
		var oAuth = new GoogleOAuth(this.oAuthName, this.oAuthScope);
		oAuth.initialize();
		//Fetch
		var url = this.oAuthScope + this.getUrlParam(min, max),
			fetchArgs = this.getFetchArgs('get', this.oAuthName),
			result,
			feed;
		try {
			result = UrlFetchApp.fetch(url, fetchArgs).getContentText(),
			Logger.log('---Throwable');
			feed = Utilities.jsonParse(result);
			Logger.log('---' + feed.data.title);
		} catch (e) {
			var serverResponse = e.toString().split('<HTML>');
			if (serverResponse.length != 2) return;

			var xml = Xml.parse('<HTML>' + serverResponse[1]);

			result = UrlFetchApp.fetch(xml.HTML.BODY.A.HREF, fetchArgs).getContentText(),
			Logger.log('---Exception');
			Logger.log('---URL:' + xml.HTML.BODY.A.HREF);
			feed = Utilities.jsonParse(result);
			Logger.log('---' + feed.data.title);

			var g_session_id = xml.HTML.BODY.A.HREF.split('gsessionid=')[1];
			ScriptProperties.setProperty('g_session_id', g_session_id);
		}
		if (!feed) return null;
		this.name = feed.data.title;
		if (!feed.data.items) return null;

		this.putInside(feed);
	},

	getFetchArgs: function(method, name) {
		return {
			method: method,
			headers: {'GData-Version': '2'},
			oAuthServiceName:name,
			oAuthUseToken:'always'
		};
	},

	getUrlParam: function(min, max) {
		var max_results = '&max-results=100',
			fields = '&title,entry(title,author,gd:when,gd:who)',
			orderby = '&orderby=starttime',
			single_events = '&singleevents=true',
			time_zone = '&ctz=Asia/Tokyo',
			start_min = '&start-min=' + min.replace('+', '%2B'),
			start_max = '&start-max=' + max.replace('+', '%2B'),
			g_session_id = '&gsessionid=' + ScriptProperties.getProperty('g_session_id'),
			urlParam = this.email + '/private/full?alt=jsonc' + max_results + fields + orderby + single_events + time_zone + start_min + start_max;
		urlParam = (g_session_id) ?urlParam + g_session_id :urlParam;
		return urlParam;
	},

	putInside: function(feed) {
		var events = feed.data.items,
			name = feed.data.title;
		for (var i = 0; i < events.length; i++) {
			this.events.push(new ClassEvent(events[i], name));
		}
	},

	getEvents: function() {
		return this.events;
	}
};

var ClassEvent = function(event, name) {
	this.name = name;
	this.event = event;
};
ClassEvent.prototype = {
	name: undefined,

	event: undefined,

	getCalendarName: function() {
		return this.name;
	},

	getCreator: function() {
		return this.event.creator;
	},

	getDateCreated: function() {
		var createdW3cdtf = this.event.created.replace('.000',''),
			createdDate = new Date();
		createdDate.setW3CDTF(createdW3cdtf);
		return createdDate;
	},

	getTitle: function() {
		return this.event.title;
	},

	getStatus: function(email) {
		var attendees = this.event.attendees;
		for (var i = 0; i < attendees.length; i++) {
			var attendee = attendees[i];
			if (attendee.email == email) {
				var status = '';
				switch (attendee.status) {
					case 'invited':
						status = '承認待ち';
						break;
					case 'declined':
						status = '辞退';
						break;
					case 'accepted':
						status = '承諾';
						break;
				}
				return status;
			}
		}
		return null;
	},

	getStartTime: function() {
		var startW3cdtf = this.event.when[0].start.replace('.000',''),
			startDate = new Date();
		if (startW3cdtf.length == 10) {
			var year = startW3cdtf.substr(0, 4),
				month = startW3cdtf.substr(5, 2) - 1,
				date = startW3cdtf.substr(8, 2);
			startDate = new Date(year, month, date);
		} else {
			startDate.setW3CDTF(startW3cdtf);
		}
		return startDate;
	},

	getEndTime: function() {
		var endW3cdtf = this.event.when[0].end.replace('.000',''),
			endDate = new Date();
		if (endW3cdtf.length == 10) {
			var year = endW3cdtf.substr(0, 4),
				month = endW3cdtf.substr(5, 2) - 1,
				date = endW3cdtf.substr(8, 2);
			endDate = new Date(year, month, date);
		} else {
			endDate.setW3CDTF(endW3cdtf);
		}
		return endDate;
	}
};

var OverBookingResearcher = function() {
	this.config = new Config();
	this.executePlace = ScriptProperties.getProperty('execute_place');
	if (!this.executePlace) this.executePlace = 0;
	this.resourcesList = ScriptProperties.getProperty('resourcesList_' + String(Math.floor(this.executePlace))).split(',');
};
OverBookingResearcher.prototype = {
	config: undefined,

	executePlace: null,

	resourcesList: null,

	setExecutePlace: function() {
		var articles = ScriptProperties.getProperty('resourcesList_articles');
		articles = Number(articles) - 1;
		if (articles != this.executePlace) {
			var execute_num = Number(this.executePlace) + 1;
			ScriptProperties.setProperty('execute_place', String(execute_num));
		} else {
			ScriptProperties.setProperty('execute_place', '0');
		}
	},

	main: function() {
		this.setExecutePlace();
		var scriptStTime = new Date();
		Logger.log('ScriptStTime:' + scriptStTime.toLocaleString());

		var bookingEvents = this.checkEvent(),
			length = bookingEvents.length;

		if (bookingEvents.length != 0) {
			this.alertMailSending(this.config.ALERT_MAIL_ADDRESS, bookingEvents, length, true);
			if (this.config.IS_SEND_MAIL_CREATOR) {
				var creators = this.divideCreators(bookingEvents, length);
				for (var i = 0; i < creators.length; i++) {
					var creator = creators[i][0],
						events = creators[i][1];
					this.alertMailSending(creator, events, events.length, false);
				}
			}
		}

		var scriptEnTime = new Date();
		Logger.log('ScriptEnTime: ' + scriptEnTime.toLocaleString());
	},

	getDivideEvents: function(allEvents) {
		var events = [],
			j = 0;
		events[j] = [];
		for (var i = 0; i < allEvents.length; i++) {
			var evt1 = allEvents[i], evt2 = allEvents[i + 1],
				d1 = (evt1) ?evt1.getStartTime() :undefined,
				d2 = (evt2) ?evt2.getStartTime() :undefined,
				dd1 = (d1) ?d1.getW3CDTF().split('T')[0] : undefined,
				dd2 = (d2) ?d2.getW3CDTF().split('T')[0] : undefined;
			events[j].push(allEvents[i]);
			if (dd1 != dd2) {
				j++;
				events[j] = [];
			}
		}
		return events;
	},

	checkEvent: function() {
		if (!this.resourcesList) return;

		var now1 = new Date(),
			now2 = new Date(),
			eventLog = 0,
			repeatLog = 0,
			bookingEvts = [];
		now2.setTime(now1.getTime() + (Number(this.config.MAX_DATE) * 24 * 3600 * 1000));
		var start_min = now1.getW3CDTF(),
			start_max = now2.getW3CDTF();

		Logger.log('');
		Logger.log('ChekingFrom: ' + start_min);
		Logger.log('ChekingTo: ' + start_max);
		Logger.log('');

		Logger.log(String(this.executePlace) + ' of resources');
		var resources = this.resourcesList;
		Logger.log('');
		Logger.log('---resources: ' + resources.length);
		for (var k = 0; k < resources.length; k++) {
			var resourceId = resources[k],
				cal = new ClassCalendar(resourceId);
			cal.fetch(start_min, start_max);
			var allEvts = cal.getEvents();

			if (!allEvts) continue;

			var divideEvts = this.getDivideEvents(allEvts);
			for (var intDay = 0; intDay < divideEvts.length; intDay++) {
				var evts = divideEvts[intDay];
					length = evts.length;
				eventLog += length;
				if (length > 0) {
					for (var i = 0; i < length; i++) {
						for(var j = i + 1; j < length; j++) {
							if (i == j) continue;
							repeatLog++;
							var evt1 = evts[i],
								evt2 = evts[j];

							//evt1 get start and end
							var evt1StTime = evt1.getStartTime(),
								evt1EnTime = evt1.getEndTime(),
								evt1_status = evt1.getStatus(resourceId);

							//evt2 get start and end
							var evt2StTime = evt2.getStartTime(),
								evt2EnTime = evt2.getEndTime(),
								evt2_status = evt2.getStatus(resourceId);

							if (evt1_status == '辞退' || evt2_status == '辞退') continue;

							//Check1 - evt2がevt1の開始時刻より開始が早く、終了時刻がevt1の開始時刻よりも遅い
							if ((evt1StTime.getTime() >= evt2StTime.getTime()) && (evt1StTime.getTime() < evt2EnTime.getTime())) {
								bookingEvts.push([evt1, evt2, resourceId, cal.name]);
								continue;
							}
							//Check2 - evt2がevt1の開始時刻より開始が遅く、開始時刻がevt1の終了時刻よりも早い
							if ((evt1StTime.getTime() < evt2StTime.getTime()) && (evt1EnTime.getTime() > evt2StTime.getTime())) {
								bookingEvts.push([evt1, evt2, resourceId, cal.name]);
								continue;
							}
						}
					}
				}
			}
		}
		Logger.log('');
		Logger.log('----events: ' + eventLog);
		Logger.log('----repeat: ' + repeatLog);
		return bookingEvts;
	},

	divideCreators: function(bookingEvents, length) {
		var srhCreators = []
			creators = [];
		for (var i = 0; i < length; i++) {
			var evt1 = bookingEvents[i][0],
				evt2 = bookingEvents[i][1],
				creator1 = evt1.getCreator().email,
				creator2 = evt2.getCreator().email;
			if (!srhCreators.contains(creator1)) srhCreators.push(creator1);
			if (!srhCreators.contains(creator2)) srhCreators.push(creator2);
		}
		for (var i = 0; i < srhCreators.length; i++) {
			var events = [];
			for (var j = 0; j < length; j++) {
				var evt1 = bookingEvents[j][0],
					evt2 = bookingEvents[j][1],
					resourceId = bookingEvents[j][2],
					resourceName = bookingEvents[j][3],
					creator1 = evt1.getCreator().email,
					creator2 = evt2.getCreator().email;
				if (srhCreators[i] == creator1 || srhCreators[i] == creator2) {
					events.push([evt1, evt2, resourceId, resourceName]);
				}
			}
			creators.push([srhCreators[i], events]);
		}
		//Log
		Logger.log('-sendingMail creator Info');
		for (var i = 0; i < creators.length; i++) {
			Logger.log('---creator: ' + creators[i][0]);
			Logger.log('-----events: ' + creators[i][1].length);
		}
		return creators;
	},

	alertMailSending: function(address, bookingEvents, length, isAdmin) {
		Logger.log('MailTo: ' + address);

		var body = ['\n'];
		for (var i = 0; i < length; i++) {
			var evt1 = bookingEvents[i][0],
				evt2 = bookingEvents[i][1],
				resourceId = bookingEvents[i][2],
				resourceName = bookingEvents[i][3],
				evt1_title = '「' + evt1.getTitle() + '」',
				evt2_title = '「' + evt2.getTitle() + '」';
			if (!isAdmin && address != evt1.getCreator().email) evt1_title = '';
			if (!isAdmin && address != evt2.getCreator().email) evt2_title = '';
			body = body.concat([
				'【' + resourceName + '】',
				'\n',
				'[',
				'登録者：' + evt1.getCreator().email + '  ',
				'登録日：' + this.convertDate(evt1.getDateCreated()),
				']',
				'\n',
				this.getEventDate(evt1),
				' :',
				evt1_title,
				'（' + evt1.getStatus(resourceId) + '）',
				'\n',
				'[',
				'登録者：' + evt2.getCreator().email + '  ',
				'登録日：' + this.convertDate(evt2.getDateCreated()),
				']',
				'\n',
				this.getEventDate(evt2),
				' :',
				evt2_title,
				'（' + evt2.getStatus(resourceId) + '）',
				'\n\n',
			]);
		}
		MailApp.sendEmail(
			//To
			address,
			//Title
			'予定重複のお知らせ',
			//Body
			'施設カレンダーにおいて予定の重複が確認されたため、このアラートメールを送信しています。\n' +
			'ご予定の確認をして頂けるよう、お願いいたします。\n' +
			body.join(''),
			{noReply: 'True'}
		);
	},

	getEventDate: function(event) {
		var start = event.getStartTime(),
			end = event.getEndTime(),
			date;
		date = [
			this.convertDate(start),
			' 〜 ',
			this.convertDate(end)
		];
		return date.join('');
	},

	convertDate: function(date) {
		var week = new Array('(日)', '(月)', '(火)', '(水)', '(木)', '(金)', '(土)'),
			cDate;
		cDate = [
			date.getFullYear(),
			'/',
			this.dateZeroFill(date.getMonth() + 1),
			'/',
			this.dateZeroFill(date.getDate()),
			week[date.getDay()],
			' ',
			this.dateZeroFill(date.getHours()),
			':',
			this.dateZeroFill(date.getMinutes())
		];
		return cDate.join('');
	},

	dateZeroFill: function(num) {
		return String('0' + String(num)).slice(-2)
	}
};
Array.prototype.contains = function(value) {
	for (var i in this) {
		if (this.hasOwnProperty(i) && this[i] === value) {
			return true;
		}
	}
	return false;
};
Date.prototype.setW3CDTF = function(dtf) {
	var sp = dtf.split(/[^0-9]/);
	if (sp.length < 6 || sp.length > 8) return;

	if (sp.length == 7) {
		if (dtf.charAt(dtf.length-1) != "Z") return;
	}
	// to numeric
	for (var i=0; i<sp.length; i++) sp[i] = sp[i]-0;
	// invalid date
	if (sp[0] < 1970 ||           // year
		sp[1] < 1 || sp[1] > 12 ||// month
		sp[2] < 1 || sp[2] > 31 ||// day
		sp[3] < 0 || sp[3] > 23 ||// hour
		sp[4] < 0 || sp[4] > 59 ||// min
		sp[5] < 0 || sp[5] > 60 ) // sec
	{
		return;
	}
	// get UTC milli seconds
	var msec = Date.UTC(sp[0], sp[1]-1, sp[2], sp[3], sp[4], sp[5]);
	// time zene offset
	if (sp.length == 8) {
		if (dtf.indexOf("+") < 0) sp[6] *= -1;
		if (sp[6] < -12 || sp[6] > 13) return;// time zone offset hour
		if (sp[7] < 0 || sp[7] > 59) return;  // time zone offset min
		msec -= (sp[6]*60+sp[7]) * 60000;
	}
	// set by milli second;
	return this.setTime(msec);
};
Date.prototype.getW3CDTF = function() {
	var year = this.getFullYear(),
		mon  = this.getMonth() + 1,
		day  = this.getDate(),
		hour = this.getHours(),
		min  = this.getMinutes(),
		sec  = this.getSeconds();
	if (mon  < 10) mon  = "0" + mon;
	if (day  < 10) day  = "0" + day;
	if (hour < 10) hour = "0" + hour;
	if (min  < 10) min  = "0" + min;
	if (sec  < 10) sec  = "0" + sec;

	var tzos = this.getTimezoneOffset(),
		tzhour = tzos / 60,
		tzmin  = tzos % 60,
		tzpm = ( tzhour > 0 ) ? "-" : "+";
	if ( tzhour < 0 ) tzhour *= -1;
	if ( tzhour < 10 ) tzhour = "0" + tzhour;
	if ( tzmin	< 10 ) tzmin  = "0" + tzmin;

	var dtf = year+"-"+mon+"-"+day+"T"+hour+":"+min+":"+sec+tzpm+tzhour+":"+tzmin;
	return dtf;
};

function settings() {
	var Ui = new SettingUiApp();
	Ui.configuration();
};

function getResource() {
	var Store = new CalendarResource();
	Store.setProperty();
};

function checkEvent() {
	var Researcher = new OverBookingResearcher();
	Researcher.main();
};

function onOpen() {
	var menus = [
		{name: 'Settings', functionName: 'settings'},
		{name: 'Get Resource', functionName: 'getResource'},
		{name: 'Check Event', functionName: 'checkEvent'}
	];

	var ss = SpreadsheetApp.getActiveSpreadsheet();
	ss.addMenu('rakumo', menus);
};

function onClick(e) {
	var userAddress = Session.getUser().getEmail(),
		domain = userAddress.split('@'),
		date = e.parameter.max_date,
		address = e.parameter.address,
		isSend = e.parameter.is_send,
		app = UiApp.getActiveApplication();

	if (date > 7) {
		app.getElementById('label').setText('調査範囲の上限は７日です');
		return app;
	}
	if (!domain || !date || !address) {
		app.getElementById('label').setText('入力されていない箇所があります');
		return app;
	}

	ScriptProperties.setProperty('DOMAIN', domain[1]);
	ScriptProperties.setProperty('MAX_DATE', date);
	ScriptProperties.setProperty('ALERT_MAIL_ADDRESS', address);
	ScriptProperties.setProperty('IS_SEND_MAIL_CREATOR', isSend);

	Browser.msgBox('設定されました。');
	return onClick;
};
