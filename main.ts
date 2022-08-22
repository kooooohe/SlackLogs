// TODO セル数をカウントして、シートを増やす sheet_name_${num}的な感じで
type Message = {
	type: string
	user: string
	text: string
	ts: string
	thread_ts?: string
	reply_count?: number
	files?: File[]
	// original
	fileURLs?: string[] // Drive's URL
}

type File = {
	id: string
	name: string
	url_private_download: string
}



const Run = () => {
	const TS_COlUMN = 7

	const header = [
	/*1*/	'date',
	/*2*/	'user',
	/*3*/	'text',
	/*4*/	'fileURLs',
	/*5*/	'replyCount',
	/*6*/	'isThread',
	/*7*/	'ts' 
	]

	const FOLDER_NAME = "SlackLogs";

	const folderID = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
	if (!folderID) {
		throw 'no set folderID';
	}
	const token = PropertiesService.getScriptProperties().getProperty('SLACK_TOKEN');
	if (!token) {
		throw 'no set slack token';
	}
	let folder = DriveApp.getFolderById(folderID)

	const fitr = folder.getFoldersByName(FOLDER_NAME);
	if (fitr.hasNext()) {
		folder =  fitr.next();
	} else {
		folder = folder.createFolder(FOLDER_NAME).setName(FOLDER_NAME)
	}


	const slack = new SlackApp(token)
	const members = slack.Members()
	console.log(members)
	const cs = slack.Channels()
	console.log(cs)



	// make all sheet by names
	// TODO: make spread sheet for each
	let ts:GoogleAppsScript.Spreadsheet.Spreadsheet

	// TODO: heavy at firsttime
	console.time('a')

	for (let c of cs) {
		const fit = folder.getFilesByName(c.name)
		if (fit.hasNext()) {
			const f = fit.next()
			ts = SpreadsheetApp.open(f)

		} else {
			ts = SpreadsheetApp.create(c.name);
			folder.addFile(DriveApp.getFileById(ts.getId()));
			// join channels if there is no sheet with the channel's name
			slack.Join(c.id)
			console.log("create")
		}

		if (ts.getSheetByName(c.name) === null) {
			 ts.insertSheet(c.name)
		}


		const ss2 = new SpreadSheetHandler(ts)

		const lastTs = ss2.LastRowCell(c.name, TS_COlUMN)
		const isFirst = lastTs === ''

		let ms = slack.Messages(c.id, lastTs)


		const foit = folder.getFoldersByName(c.name)
		// for Download
		let tFolder:GoogleAppsScript.Drive.Folder 
		if (foit.hasNext()) {
			tFolder = foit.next()
		} else {
			tFolder = folder.createFolder(c.name)
		}

		ms = downloadFiles(ms,slack,tFolder)


		let svs = formatToTwoDimentions(ms, members)
		if (isFirst) {
			svs.unshift(header)
		} else {
			// it gets one same msg	
			svs.shift()
		}

		console.log(c)
		ss2.SetValues(c.name, svs)
	}
	console.timeEnd('a')
}

const downloadFiles = (ms: Message[], slack: SlackApp, tFolder: GoogleAppsScript.Drive.Folder): Message[] =>  {
	for (let m of ms) {
		if (m.files) {
			for (let f of m.files) {
				const data = slack.Download(f.url_private_download)
				const fname = f.name + '_' + f.name
				const b = data.getBlob().setName(fname)
				const fit = tFolder.getFilesByName(fname)
				let tFile:GoogleAppsScript.Drive.File
				if (fit.hasNext()) {
					tFile = fit.next()
				} else {
					tFile = tFolder.createFile(b)
				} 
				if (m.fileURLs) {
					m.fileURLs.push(tFile.getUrl())
				} else {
					m.fileURLs = [tFile.getUrl()]
				}
			}
		}
	}
	return ms
}

const formatToTwoDimentions = (ms: Message[], members: object): string[][] => {
	let tmpss:string[][] = []
	for (let m of ms) {
		//date, user, text, fileURLs, reply_count, isThread,ts 
		const utime = m.ts.substring(0, m.ts.indexOf("."))

		const d = new Date(Number(utime) * 1000)
		let rc = '0'
		if  (m.reply_count) {
			rc = String(m.reply_count)
		}

		let fileURLS = ''
		if (m.fileURLs) {
			fileURLS = m.fileURLs.join("\n")
		}

		let threadTs = '0'
		if (m.thread_ts) {
			threadTs = '1'
		}

		let text = m.text
		for (let k in members) {
			text = text.replace(k, members[k])		
		}

		const tmp:string[] = [
			d.toString(),
			members[m.user],
			text,
			fileURLS,
			rc,
			threadTs,
			m.ts
		]

		tmpss.push(tmp)
	}

	return  tmpss
}



class SpreadSheetHandler {
	private ss: GoogleAppsScript.Spreadsheet.Spreadsheet

	 // constructor(folder:GoogleAppsScript.Drive.Folder, fileName:string) {
	 constructor(s: GoogleAppsScript.Spreadsheet.Spreadsheet) {
		 this.ss = s
		
		 // TODO: delete check
		// const it = folder.getFilesByName(fileName);
		// if (it.hasNext()) {
		// 	const file = it.next();
		// 	this.ss = SpreadsheetApp.openById(file.getId());
		// }
		// else {
		// 	const ss = SpreadsheetApp.create(fileName);
		// 	folder.addFile(DriveApp.getFileById(ss.getId()));
		// 	this.ss = ss
		// }
	}

	public SetValues(sheetName:string, vs: string[][]) {
		console.log(vs)
		if (vs.length === 0) {
			return
		}
		const ts = this.ss.getSheetByName(sheetName)
		const lastRow = ts.getLastRow()
		const startRow = lastRow+1
		ts.getRange(startRow, 1, vs.length, vs[0].length ).setValues(vs)
	}

	public Utlity() {
		return this.ss
	}

	// to get last ts
	public LastRowCell(sheetName:string, column: number) {
		const ts = this.ss.getSheetByName(sheetName)
		const ll = ts.getLastRow()
		if (ll === 0) {
			return ''
		}
		return ts.getRange(ll,column).getValue()
	}

}


class SlackApp {
	private token: string = '';
	private baseURL = "https://slack.com/api/";
	// "https://slack.com/api/" + path + "?";
	private REQUEST_MESSAGE_LIMIT = 5
	private REQUEST_THREAD_LIMIT = 5
	private messageRequesstCount = 0
	private threadRequesstCount = 0

	constructor (token:string) {
		this.token = token;
	}

	private request (path: string, params:object):any {
		let url = this.baseURL + path + "?";
		let queries = [];
		for (let k in params) {
			queries.push(encodeURIComponent(k) + "=" + encodeURIComponent(params[k]));
		}
		url += queries.join('&');
		const headers = {
			'Authorization': 'Bearer ' + this.token
		};

		const options = {
			'headers': headers, 
		};
		let data:any
		try {
			const response = UrlFetchApp.fetch(url, options);
			data = JSON.parse(response.getContentText());
		} catch(error) {
			console.log(error)
			throw error
		}
		return data;
	};

	public Members () {
		const data = this.request('users.list', {});
		let members = new Map<string,string>
		for (let m of data.members) {
			members[m.id] = m.name;
		}
		return members;
	};
	public Channels () {
		type Channel  = {
			id: string
			name: string
		}
		const data = this.request('conversations.list',{});
		let cs:Channel[] = []
		for (let d of data.channels) {
			const t:Channel = {id: d.id, name: d.name} 
			cs.push(t)
		}

		return cs
	};

	public Messages(channelID:string, oldest:string = '') {
		let options = new Map<string, string>()
		options['channel'] = channelID
		options['oldest'] = oldest
		let hasNext = true

		let ms: Message[] = []
		while (hasNext && this.canMessageRequest()) {
			const data = this.messages(options)
			console.log(data)
			const mstmp:Message[] = data.messages
			hasNext = !!data.has_more
			if (mstmp && mstmp.length > 0){
			  ms = ms.concat(mstmp)
			  options['oldest'] = mstmp[0].ts
			}
		}

		let nMs:Message[] = []
		// getThreads
		for (let m of ms) {
			if (m.reply_count > 0) {
				const ms = this.threadMessages(channelID, m.thread_ts)
				nMs = nMs.concat(ms.reverse())
			} else {
				nMs.push(m)	
			}
		}

		console.log(nMs)
		// let ms: Message[] = data.messages

		return nMs.reverse()
	}

	private canMessageRequest():boolean {
		this.messageRequesstCount++;
		const r = this.messageRequesstCount < this.REQUEST_MESSAGE_LIMIT 

		if (!r) {
			this.messageRequesstCount = 0
		}

		return r
	}

	private canThtradRequest():boolean {
		this.threadRequesstCount++;
		const r = this.threadRequesstCount < this.REQUEST_THREAD_LIMIT 

		if (!r) {
			this.threadRequesstCount = 0
		}
		return r
	}

	private messages(options:Map<string, string>) {
		const data = this.request('conversations.history', options)
		// const ms: Message[] = data.messages
		return data 
	}

	private threadMessages(channelID:string, threadTs:string) {
		// TODO: limit
		let options = new Map<string, string>;
		options['channel'] = channelID
		options['ts'] =  threadTs
		let hasNext = true
			console.log('get thred func1')
			console.log(channelID)

		let ms: Message[] = []
		while (hasNext && this.canThtradRequest()) {
			const data = this.request('conversations.replies', options)
			const mstmp:Message[] = data.messages
			console.log('get thred func2')
			console.log(data)
			console.log(mstmp)
			hasNext = !!data.has_more
			if (mstmp && mstmp.length > 0){
			  ms = ms.concat(mstmp)
			  options['oldest'] = mstmp[0].ts
			}
		}
			console.log('get thred func3')
			console.log(ms)

		return ms
	}

	public Join(channelID:string):void {
		let options = new Map<string, string>;
		options['channel'] = channelID
		this.request('conversations.join ', options)
	}

	public Download(url:string) {
		const headers = {
			'Authorization': 'Bearer ' + this.token
		};

		const options = {
			'headers': headers, 
		};
		let data:any
		try {
			const data = UrlFetchApp.fetch(url, options);
		} catch(error) {
			console.log(error)
			// ignore if the size is too big
		}
		return data;
	}

}

