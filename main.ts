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

const MAX_ROW = 7000
const TS_COlUMN = 7
const REPLY_COUNT_COlUMN = 5

const Run = () => {

	const header = [
		/*1*/	'date',
		/*2*/	'user',
		/*3*/	'text',
		/*4*/	'fileURLs',
		/*5*/	'replyCount',
		/*6*/	'isThread',
		/*7*/	'ts' 
	]

	const FOLDER_NAME = "SlackLog";

	const folderID = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
	if (!folderID) {
		throw 'no set folderID';
	}
	const token = PropertiesService.getScriptProperties().getProperty('SLACK_TOKEN');
	if (!token) {
		throw 'no set slack token';
	}
	let rootFolder = DriveApp.getFolderById(folderID)

	const fitr = rootFolder.getFoldersByName(FOLDER_NAME);
	if (fitr.hasNext()) {
		rootFolder =  fitr.next();
	} else {
		rootFolder = rootFolder.createFolder(FOLDER_NAME).setName(FOLDER_NAME)
	}

	const imfitr = rootFolder.getFoldersByName('img'); 
	let imgFolder:GoogleAppsScript.Drive.Folder
	if (imfitr.hasNext()) {
		imgFolder = imfitr.next()
	} else {
		imgFolder = rootFolder.createFolder('img').setName('img')
	}


	const slack = new SlackApp(token)
	const members = slack.Members()
	console.log(members)
	const cs = slack.Channels()
	console.log(cs)




	// TODO: heavy at firsttime
	console.time('a')

	for (let c of cs) {
		let ts:GoogleAppsScript.Spreadsheet.Spreadsheet
		const fit = rootFolder.getFilesByName(c.name)
		if (fit.hasNext()) {
			const f = fit.next()
			ts = SpreadsheetApp.open(f)

		} else {
			ts = SpreadsheetApp.create(c.name);
			rootFolder.addFile(DriveApp.getFileById(ts.getId()));
			// join channels if there is no sheet with the channel's name
			slack.Join(c.id)
			console.log("create")
		}

		let ss2 = new SpreadSheetHandler(ts, c.name)

		if (ss2.IsOverMaxRow()) {
			const d = new Date()
			const y = d.getFullYear();
			const m = ('00' + (d.getMonth()+1)).slice(-2);
			const da = ('00' + d.getDate()).slice(-2);
			const r =  (y + '_' + m + '_' + da);
			ss2.Utlity().rename(ss2.sheetName+'_'+r)

			ts = SpreadsheetApp.create(c.name);
			rootFolder.addFile(DriveApp.getFileById(ts.getId()));

			ss2 = new SpreadSheetHandler(ts, c.name)
		}



		const lastTs = ss2.LastRowCell(TS_COlUMN, 1)

		let ms = slack.Messages(c.id, lastTs)


		const foit = imgFolder.getFoldersByName(c.name)
		// for Download
		let tFolder:GoogleAppsScript.Drive.Folder 
		if (foit.hasNext()) {
			tFolder = foit.next()
		} else {
			tFolder = imgFolder.createFolder(c.name)
		}

		ms = downloadFiles(ms,slack,tFolder)

		const isFirst = ss2.IsNothing()

		let svs = formatToTwoDimentions(ms, members)
		if (isFirst) {
			svs.unshift(header)
		} else {
			// it gets one same msg	
			if (lastTs) {
			  svs.shift()
			}
		}

		console.log(c)
		ss2.SetValues(svs)

		const tms = ss2.Threads()
		for (let tm of tms) {
			let ms = slack.ThreadMessages(c.id, tm.threadTs, tm.lastTs)
			ms = downloadFiles(ms,slack,tFolder)
			let svs = formatToTwoDimentions(ms, members)
			// it gets one same msg	
			svs.shift()
			ss2.Insert(svs, tm.rowNum)
			
		}
	}
	console.timeEnd('a')
}

const downloadFiles = (ms: Message[], slack: SlackApp, tFolder: GoogleAppsScript.Drive.Folder): Message[] =>  {
	for (let m of ms) {
		if (m.files) {
			for (let f of m.files) {
				let data:any
				try {
					data = slack.Download(f.url_private_download)
				} catch(error) {
					// for deleteed file
					console.log(error)
					continue
				}
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


class FolderHandler {
	private f: GoogleAppsScript.Drive.Folder

	constructor(f: GoogleAppsScript.Drive.Folder) {
		this.f = f
	}
}

class SpreadSheetHandler {
	private ss: GoogleAppsScript.Spreadsheet.Spreadsheet
	public sheetName: string
	private TREAD_TARGET_RANGE = 500


	// constructor(folder:GoogleAppsScript.Drive.Folder, fileName:string) {
	constructor(s: GoogleAppsScript.Spreadsheet.Spreadsheet, sheetName: string) {
		this.ss = s
		this.sheetName = sheetName
		if (this.ss.getSheetByName(this.sheetName) === null) {
			this.ss.insertSheet(this.sheetName)
		}
	}

	public SetValues(vs: string[][]) {
		console.log(vs)
		if (vs.length === 0) {
			return
		}
		const ts = this.ss.getSheetByName(this.sheetName)
		const lastRow = ts.getLastRow()
		const startRow = lastRow+1
		ts.getRange(startRow, 1, vs.length, vs[0].length).setNumberFormat('@').setValues(vs)
	}

	public Utlity() {
		return this.ss
	}

	// to get last ts
	public LastRowCell(column: number, ignore:number = 0) {
		const ts = this.ss.getSheetByName(this.sheetName)
		const ll = ts.getLastRow()
		if (ll <= ignore) {
			return ''
		}
		return ts.getRange(ll,column).getValue()
	}

	public IsNothing() {
		const ts = this.ss.getSheetByName(this.sheetName)
		const ll = ts.getLastRow()
		return ll === 0
	}

	public IsOverMaxRow() {
		const ts = this.ss.getSheetByName(this.sheetName)
		return ts.getLastRow() > MAX_ROW
	}

	public Threads() {
		const ts = this.ss.getSheetByName(this.sheetName)	
		const lastRow = ts.getLastRow()
		let startRow = lastRow
		if (startRow < this.TREAD_TARGET_RANGE) {
			startRow = 2 // without header
		} else {
			startRow -= this.TREAD_TARGET_RANGE
			startRow++
		}

		if (lastRow < startRow) {
			return []
		}

		const trs = ts.getRange(startRow,REPLY_COUNT_COlUMN,lastRow-1,3).getValues()
		//replyCount,isThread,ts
		console.log(trs)
		type LastThredMessages = {
			threadTs: string
			lastTs: string
			rowNum: number
		}

		const REPLY = 0
		const IS_THREAD = 1
		const TS = 2
		let r:LastThredMessages[] = []
		let p = [0,0,'']
		let lastThreadTs = ''
		let cnt = 0
		for (let t of trs) {
			if (t[REPLY] != 0) {
				lastThreadTs = String(t[TS])
			}

			if (t[IS_THREAD] == 0) {
				if(p[IS_THREAD] == 1) {
					const tr:LastThredMessages = {
						lastTs: String(p[TS]),
						threadTs: lastThreadTs,
						rowNum: startRow + cnt

					}
					r.push(tr)	
				}
			}

			p = t
			cnt++
		}
		return r
	}

	public Insert(vs:string[][], rowNum:number) {
		if (vs.length === 0) {
			return
		}
		const ts = this.ss.getSheetByName(this.sheetName)
		ts.insertRowsAfter(rowNum - 1, vs.length)
		// const lastRow = ts.getLastRow()
		ts.getRange(rowNum, 1, vs.length, vs[0].length).setNumberFormat('@').setValues(vs)
	}
}


class SlackApp {
	private token: string = '';
	private baseURL = "https://slack.com/api/";
	// "https://slack.com/api/" + path + "?";
	private REQUEST_MESSAGE_LIMIT = 5
	private REQUEST_THREAD_LIMIT = 3
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
			if(m.profile?.display_name) {
				members[m.id] = m.profile.display_name;
				continue;
			}

			if(m.real_name) {
				members[m.id] = m.real_name;
				continue;
			}
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

	public ThreadMessages(channelID:string, threadTs:string, oldest:string) {
		const ms = this.threadMessages(channelID,threadTs,oldest)
		return ms
	}
	private threadMessages(channelID:string, threadTs:string, oldest:string = '') {
		let options = new Map<string, string>;
		options['channel'] = channelID
		options['ts'] =  threadTs
		options['oldest'] = oldest
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
			throw(error)
			// ignore if the size is too big
		}
		return data;
	}

}

