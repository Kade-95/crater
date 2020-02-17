import { ElementModifier, func, ColorPicker, Connection, BaseClientSideWebPart } from '.';
import FroalaEditor from 'froala-editor';
import 'froala-editor/js/plugins/align.min.js';

require('./../styles/connection.css');
require('./../styles/form.css');
require('./../styles/events.css');
require('./../styles/event.css');
require('./../styles/list.css');
require('./../styles/slider.css');
require('./../styles/counter.css');
require('./../styles/tiles.css');
require('./../styles/sections.css');
require('./../styles/news.css');
require('./../styles/table.css');
require('./../styles/panel.css');
require('./../styles/special.css');
require('./../styles/textarea.css');
require('./../styles/icons.css');
require('./../styles/button.css');
require('./../styles/countdown.css');
require('./../styles/tab.css');
require('./../styles/carousel.css');
require('./../styles/map.css');
require('./../styles/datelist.css');
require('./../styles/instagram.css');
require('./../styles/beforeafter.css');
require('./../styles/table.css');
require('./../styles/power.css');
require('./../styles/employeedirectory.css');
require('./../styles/birthday.css');
require('./../../../node_modules/froala-editor/css/froala_editor.pkgd.min.css');
const factory = require('./powerbi.js');

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

class CraterWebParts {
	public elementModifier = new ElementModifier();
	public sharePoint: any;
	public connectable: any = [
		'list', 'slider', 'counter', 'tiles', 'news', 'table', 'icons', 'button', 'events', 'carousel', 'datelist', 'event'
	];

	constructor(params) {
		this.sharePoint = params.sharePoint;
	}
	//create the pane-options component
	public paneOptions(params) {
		//create the options element
		let options = this.elementModifier.createElement({
			element: 'span', attributes: { class: 'crater-content-options', style: { visibility: 'hidden' } }
		});

		//set the options data
		let paneOptionsData = {
			'AB': { title: 'Add Before', class: 'add-before' },
			'AA': { title: 'Add After', class: 'add-after' },
			'D': { title: 'Delete', class: 'delete' }
		};

		//append the options data to options
		if (func.isset(params.options)) {
			for (let option of params.options) {
				options.makeElement({
					element: 'button', text: option, attributes: { class: `${paneOptionsData[option].class}-${params.owner} small btn`, title: paneOptionsData[option].title }
				});
			}
		}
		else {
			options.append(
				this.elementModifier.createElement({
					element: 'button', text: 'AB', attributes: { class: `add-before-${params.owner} small btn`, title: 'Add Before' }
				}),
				this.elementModifier.createElement({
					element: 'button', text: 'AA', attributes: { class: `add-after-${params.owner} small btn`, title: 'Add After' }
				}),
				this.elementModifier.createElement({
					element: 'button', text: 'D', attributes: { class: `delete-${params.owner} small btn`, title: 'Delete Row' }
				})
			);
		}

		return options;
	}

	// setup the color picker compnent
	public pickColor(params, callBack) {
		//remove all color pickers
		params.parent.findAll('.pick-color').forEach(element => {
			element.remove();
		});
		this.elementModifier.sharepoint = this.sharePoint;

		params.parent.addEventListener('mouseenter', (event) => {
			//on hover add color picker

			let options = params.parent.makeElement({
				//set the options
				element: 'img', attributes: { class: 'small btn pick-color', src: this.sharePoint.images.edit, style: { cursor: 'pointer', backgroundColor: 'white', width: '1em', height: 'auto', position: 'absolute', top: '0px', right: '0px' } }
			});
			params.parent.css({ position: 'relative' });

			options.addEventListener('click', () => {
				//pick the color
				params.parent.findAll('.crater-color-picker').forEach(element => {
					element.remove();
				});
				let colorPicker = new ColorPicker({ width: 200, height: 200 });
				params.parent.makeElement({ element: colorPicker.canvas });
				colorPicker.draw(0.1);
				colorPicker.onChanged(callBack);
			});
		});

		params.cell.onChanged(callBack);

		params.parent.addEventListener('mouseleave', (event) => {
			params.parent.findAll('.pick-color').forEach(element => {
				element.remove();
			});
		});
	}

	//set up the image uploader
	public uploadImage(params, callBack) {
		//remove all the uploader
		params.parent.findAll('.upload-form').forEach(element => {
			element.remove();
		});
		this.elementModifier.sharepoint = this.sharePoint;

		params.parent.addEventListener('mouseenter', (event) => {
			//on hover add the option
			let add = params.parent.makeElement({
				element: 'img', attributes: { class: 'small btn upload-image', src: this.sharePoint.images.edit, style: { cursor: 'pointer', backgroundColor: 'white', width: '1em', height: 'auto', position: 'absolute', top: '0px', right: '0px' } }
			});
			params.parent.css({ position: 'relative' });

			add.addEventListener('click', () => {
				params.parent.findAll('.upload-form').forEach(element => {
					element.remove();
				});
				//add the uploader
				this.elementModifier.importImage({ parent: params.parent, name: 'upload', attributes: { class: 'upload-form' } }, (image) => {
					params.parent.findAll('.upload-form').forEach(element => {
						element.remove();
					});
					callBack(image);
				});
			});
		});

		params.parent.addEventListener('mouseleave', (event) => {
			params.parent.findAll('.upload-image').forEach(element => {
				element.remove();
			});
		});
	}

	//create the webpart options
	public webPartOptions(params) {
		let optionContainer = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'webpart-options', style: {} }
		});

		let options = {
			append: this.sharePoint.images.append,
			view: this.sharePoint.images.view,
			edit: this.sharePoint.images.edit,
			delete: this.sharePoint.images.delete,
			clone: this.sharePoint.images.copy,
			paste: this.sharePoint.images.paste,
			undo: this.sharePoint.images.undo,
			redo: this.sharePoint.images.redo
		};

		for (let option of params.options) {
			optionContainer.makeElement({
				element: 'img', attributes: {
					class: 'webpart-option', id: option.toLowerCase() + '-me', src: options[option.toLowerCase()], alt: option, title: `${option} ${params.title}`, style: { display: option.toLowerCase() == 'paste' ? 'none' : '' }
				}
			});
		}

		if (func.isset(params.attributes)) optionContainer.setAttributes(params.attributes);
		return optionContainer;
	}

	//show webpart options
	public showOptions(element) {

		let handlePaste = (webpart) => {
			if (webpart.classList.contains('crater-container')) {
				if (this.sharePoint.pasteActive) {
					webpart.find('.webpart-options').find('#paste-me').show();
				} else {
					webpart.find('.webpart-options').find('#paste-me').hide();
				}
			}
		};

		element.addEventListener('mouseenter', event => {
			this.sharePoint.dontSave = true;
			if (element.hasAttribute('data-key') && this.sharePoint.inEditMode()) {
				element.find('.webpart-options').show();
				handlePaste(element);
			}
		});

		element.addEventListener('mouseleave', event => {
			this.sharePoint.dontSave = true;
			if (element.hasAttribute('data-key')) {
				element.find('.webpart-options').hide();
			}
		});

		element.findAll('.keyed-element').forEach(keyedElement => {
			keyedElement.addEventListener('mouseenter', event => {
				this.sharePoint.dontSave = true;
				if (keyedElement.hasAttribute('data-key') && this.sharePoint.inEditMode()) {
					keyedElement.find('.webpart-options').show();
					handlePaste(keyedElement);
				}
			});

			keyedElement.addEventListener('mouseleave', event => {
				this.sharePoint.dontSave = true;
				if (keyedElement.hasAttribute('data-key')) {
					keyedElement.find('.webpart-options').hide();
				}
			});
		});
	}

	//generate webpart key
	public generateKey() {
		let found = true;
		let key = func.generateRandom(10);
		while (found) {
			key = func.generateRandom(10);
			found = this.sharePoint.properties.pane.content.hasOwnProperty(key);
		}
		return key;
	}

	//create webpart element
	public createKeyedElement(params) {
		let key = this.generateKey();
		if (!func.isset(params.attributes)) params.attributes = {};
		params.attributes['data-key'] = key;
		if (!func.isset(params.attributes['data-type'])) params.attributes['data-type'] = 'sample';

		this.sharePoint.properties.pane.content[key] = { content: '', styles: '', connection: '', settings: {}, sync: {}, draft: { dom: '', html: '', pane: { content: '', styles: '', connection: '' } } };

		if (!func.isset(params.options)) params.options = ['Edit', 'Delete', 'Clone'];

		let options = params.options;
		delete params.options;

		let element = this.elementModifier.createElement(params);

		if (element.classList.contains('crater-container')) {
			options.push('Paste');
		}

		let optionsMenu = this.webPartOptions({ options, title: params.attributes['data-type'] });

		element.prepend(optionsMenu);

		if (func.isset(params.alignOptions)) {
			let align = {};
			if (params.alignOptions == 'bottom') {
				element.find('.webpart-options').css({ top: 'unset' });
			}
			align[params.alignOptions] = '0px';
			element.find('.webpart-options').css(align);
			if (params.alignOptions == 'center') {
				element.find('.webpart-options').css({ margin: '0em 3em' });
			}
		}

		element.classList.add('keyed-element');
		if (this.connectable.includes(params.attributes['data-type'].toLowerCase())) {
			element.dataset.connectible = 'true';
		}

		return element;
	}

	//create a pane style block
	public createStyleBlock(params) {
		//create the block
		let block = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'card crater-style-block', style: { margin: '1em', position: 'relative' } }, sync: true, children: [
				this.elementModifier.createElement({
					element: 'div', attributes: { class: 'card-title' }, children: [
						this.elementModifier.createElement({
							//set the block title
							element: 'h2', attributes: { class: 'title' }, text: params.title
						})
					]
				}),
			]
		});

		let blockRow = block.makeElement({
			element: 'div', attributes: { class: 'row' }
		});

		let key = params.element.dataset['key'];

		if (func.isset(params.options)) {//set the options
			let styleOptions = block.makeElement({
				element: 'div', attributes: { class: 'crater-style-options' }
			});

			if (func.isset(params.options.sync)) {//set the sync and async option
				if (!func.isset(this.sharePoint.properties.pane.content[key].sync[params.title.toLowerCase()])) {
					this.sharePoint.properties.pane.content[key].sync[params.title.toLowerCase()] = false;
				}

				let sync = this.sharePoint.properties.pane.content[key].sync[params.title.toLowerCase()];

				styleOptions.makeElement({
					element: 'img', attributes: {
						class: 'crater-style-option crater-toggle-style-sync', alt: sync ? 'Async' : 'Sync', src: sync ? this.sharePoint.images.sync : this.sharePoint.images.async, 'data-style': params.title.toLowerCase()
					}
				});
			}
		}

		for (let i in params.children) {//append the style data to the block
			let value = '';
			if (func.isset(params.element.css()[func.cssStyleName(i)])) {
				value = params.element.css()[func.cssStyleName(i)];
			}
			let styleSync = '';
			if (func.isset(params.options) && func.isset(params.options.sync)) {
				styleSync = params.title.toLowerCase();
			}

			if (i == 'backgroundImage') {
				blockRow.append(this.elementModifier.cell({
					element: 'img', dataAttributes: { 'data-action': i, 'data-style-sync': styleSync, class: 'crater-icon crater-style-attr' }, name: params.children[i], value, src: this.sharePoint.images.append
				}));
			}
			else if (i == 'fontFamily') {
				let list = func.fontStyles;
				blockRow.append(this.elementModifier.cell({
					element: 'input', dataAttributes: { 'data-action': i, 'data-style-sync': styleSync, class: 'crater-style-attr' }, name: params.children[i], value, list
				}));
			}
			//@ts-ignore
			else if (i.includes('color') || i.includes('Color')) {
				let list = func.colors;
				blockRow.append(this.elementModifier.cell({
					element: 'input', dataAttributes: { 'data-action': i, 'data-style-sync': styleSync, class: 'crater-style-attr' }, name: params.children[i], value, list
				}));
			}
			else if (i == 'fontWeight') {
				blockRow.append(this.elementModifier.cell({
					element: 'select', dataAttributes: { 'data-action': i, 'data-style-sync': styleSync, class: 'crater-style-attr' }, name: params.children[i], value, options: func.boldness
				}));
			}
			else if (i == 'fontSize') {
				blockRow.append(this.elementModifier.cell({
					element: 'input', dataAttributes: { 'data-action': i, 'data-style-sync': styleSync, class: 'crater-style-attr' }, name: params.children[i], value, list: func.pixelSizes
				}));
			}
			else {
				blockRow.append(this.elementModifier.cell({
					element: 'input', dataAttributes: { 'data-action': i, 'data-style-sync': styleSync, class: 'crater-style-attr' }, name: params.children[i], value
				}));
			}
		}

		return block;
	}
}

class Sample extends CraterWebParts {
	constructor(params) {
		super({ sharePoint: params.sharePoint });
	}
}

class EmployeeDirectory extends CraterWebParts {
	public params: any;
	public elementModifier = new ElementModifier();
	public paneContent: any;
	public element: any;
	private key: any;
	private users: any;
	private openImage: any = 'https://pngimg.com/uploads/plus/plus_PNG22.png';
	private closeImage: any = 'https://i.dlpng.com/static/png/1442324-minus-png-minus-png-1600_1600_preview.png';

	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.params = params;
	}

	public render(params) {
		if (!func.isset(params.source)) params.source = [
			{ image: this.sharePoint.images.append, text: 'One here' },
			{ image: this.sharePoint.images.edit, text: 'Two now' },
			{ image: this.sharePoint.images.sync, text: 'Three then' },
			{ image: this.sharePoint.images.async, text: 'Four done' },
			{ image: this.sharePoint.images.delete, text: 'Five when' }
		];

		let employeeDirectory = this.createKeyedElement({
			element: 'div', attributes: { class: 'crater-employee-directory crater-component', 'data-type': 'employeedirectory' }, children: [
				{
					element: 'div', attributes: { class: 'crater-employee-directory-content' }, children: [
						{
							element: 'div', attributes: { class: 'crater-employee-directory-main-view' }, children: [
								{
									element: 'menu', attributes: { class: 'crater-employee-directory-menu' }, children: [
										{ element: 'input', attributes: { placeholder: 'Search by Name...', id: 'crater-employee-directory-search-query' } },
										{
											element: 'div', attributes: { style: { display: 'grid', justifyContent: 'center', alignContent: 'center', gridTemplateColumns: '1fr max-content' } }, children: [
												{ element: 'select', attributes: { id: 'crater-employee-directory-search-type', class: 'btn' }, options: ['All', 'By Name', 'By Department', 'By Job Title'] },
												{ element: 'img', attributes: { class: 'crater-employee-directory-icon', id: 'crater-employee-directory-sync', title: 'Refresh', src: this.sharePoint.images.sync } }
											]
										}
									]
								},
								{ element: 'div', attributes: { class: 'crater-employee-directory-display' } }
							]
						}
					]
				}
			]
		});

		this.key = this.key || employeeDirectory.dataset.key;
		let settings = this.sharePoint.properties.pane.content[this.key].settings;

		settings.searchType = 'All';
		settings.searchQuery = '';
		settings.mailApp = 'Outlook';
		settings.messageApp = 'Teams';
		settings.callApp = 'Teams';
		settings.employees = [];

		localStorage[`crater-${this.key}`] = JSON.stringify(settings);
		return employeeDirectory;
	}

	public rendered(params) {
		this.sharePoint = params.sharePoint;
		this.element = params.element;
		this.key = this.element.dataset.key;
		let displayed = false;

		let settings = this.sharePoint.properties.pane.content[this.key].settings;

		let display = this.element.find('.crater-employee-directory-display');

		let gmail = 'https://mail.google.com';
		let outlook = 'https://outlook.office365.com/mail';
		let yahoo = 'https://mail.yahoo.com';
		let skype = 'https://www.skype.com/en/business/';
		let teams = 'https://teams.microsoft.com/_#/conversations';

		display.innerHTML = '';

		display.makeElement({ element: 'img', attributes: { src: this.sharePoint.images.loading, class: 'crater-icon', style: { alignSelf: 'center', justifySelf: 'center' } } });

		let getUsers = () => {
			this.sharePoint.connection.getWithGraph().then(client => {
				client.api('/users')
					.select('mail, displayName, givenName, id, jobTitle, mobilePhone')
					.get((_error: any, _result: MicrosoftGraph.User, _rawResponse?: any) => {
						this.users = _result['value'];

						let getImage = (id) => {
							return new Promise((resolve, reject) => {
								client.api(`/users/${id}/photo/$value`)
									.responseType('blob')
									.get((error: any, result: any, rawResponse?: any) => {
										if (!func.setNotNull(result)) return;
										settings.employees[id].photo = result;
										if (displayed) {
											display.find(`#row-${id}`).find('.crater-employee-directory-dp').src = window.URL.createObjectURL(result);
										}
										resolve();
									});
							});
						};

						let getDepartment = (id) => {
							return new Promise((resolve, reject) => {
								client.api(`/users/${id}/department`)
									.get((error: any, result: any, rawResponse?: any) => {
										if (!func.setNotNull(result)) return;
										settings.employees[id].department = result.value;
										resolve();
									});
							});
						};

						for (let employee of this.users) {
							settings.employees[employee.id] = employee;
							getImage(employee.id);
							getDepartment(employee.id);
						}

						this.displayUsers(display);
						displayed = true;
					});
			});
		};

		if (!this.sharePoint.isLocal) {
			getUsers();
		}
		else {
			let sample = { mail: 'kennedy.ikeka@ipigroupng.com', id: this.key, displayName: 'Ikeka Kennedy', jobTitle: 'Programmer' };
			this.users = [];
			settings.employees = {};
			for (let i = 0; i < 100; i++) {
				this.users.push(sample);
				settings.employees[sample.id] = sample;
			}

			this.displayUsers(display);
		}

		let changeSearchType = this.element.find('#crater-employee-directory-search-type');
		let changeSearchQuery = this.element.find('#crater-employee-directory-search-query');
		let sync = this.element.find('#crater-employee-directory-sync');

		sync.addEventListener('click', event => {
			getUsers();
		});

		changeSearchType.onChanged(value => {
			settings.searchType = value;
			if (value == 'All') {
				changeSearchQuery.value = '';
				changeSearchQuery.setAttribute('value', '');
			}
			this.displayUsers(display);
		});

		changeSearchQuery.onChanged(value => {
			settings.searchQuery = value;
			this.displayUsers(display);
		});

		let menu = this.element.find('.crater-employee-directory-menu');
		if (menu.position().width < 400) {
			menu.css({ gridTemplateColumns: '1fr' });
		}
		else {
			menu.cssRemove(['grid-template-columns']);
		}

		display.addEventListener('click', event => {
			let element = event.target;

			if (element.classList.contains('crater-employee-directory-toggle-view')) {
				element.classList.toggle('open');
				let row = element.getParents('.crater-employee-directory-row');

				display.findAll('.crater-employee-directory-other-details').forEach(other => {
					other.remove();
				});
				display.findAll('.crater-employee-directory-toggle-view').forEach(toggle => {
					toggle.src = this.openImage;
				});

				if (element.classList.contains('open')) {
					element.src = this.closeImage;
					row.append(this.displayOtherDetails(settings.employees[row.id.replace('row-', '')]));
				}
			}
			else if (element.classList.contains('crater-employee-directory-mail')) {
				let row = element.getParents('.crater-employee-directory-row');

				if (settings.mailApp.toLowerCase() == 'outlook') {
					window.open(outlook);
				}
				else if (settings.mailApp.toLowerCase() == 'gmail') {
					window.open(gmail);
				}
				else if (settings.mailApp.toLowerCase() == 'yahoo') {
					window.open(yahoo);
				}
			}
			else if (element.classList.contains('crater-employee-directory-message')) {
				if (settings.messageApp.toLowerCase() == 'teams') {
					window.open(teams);
				}
				else if (settings.messageApp.toLowerCase() == 'skype') {
					window.open(skype);
				}
			}
			else if (element.classList.contains('crater-employee-directory-phone')) {
				if (settings.callApp.toLowerCase() == 'teams') {
					window.open(teams);
				}
				else if (settings.callApp.toLowerCase() == 'skype') {
					window.open(skype);
				}
			}
		});
	}

	private displayOtherDetails(user) {
		let otherDetials = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'crater-employee-directory-other-details' }, children: [
				this.elementModifier.cell({ element: 'span', text: user.department || '', name: 'Department' }),
				this.elementModifier.cell({ element: 'span', text: user.office || '', name: 'Location' }),
				this.elementModifier.cell({ element: 'span', text: user.jobTitle || '', name: 'Job' }),
				this.elementModifier.cell({ element: 'span', text: user.mobilePhone || '', name: 'Mobile Phone' })
			]
		});
		return otherDetials;
	}

	public displayUsers(display) {
		display.innerHTML = '';
		let settings = this.sharePoint.properties.pane.content[this.key].settings;
		let stored = JSON.parse(localStorage[`crater-${this.key}`]);

		for (let i = 0; i < this.users.length; i++) {
			let employee = this.users[i];
			if (settings.searchType != 'All' && settings.searchQuery != '') {

				if (settings.searchType == 'By Name') {
					if (!employee.displayName.toLowerCase().includes(settings.searchQuery.toLowerCase())) continue;
				}
				else if (settings.searchType == 'By Department') {
					let department = settings.employees[employee.id].department || stored.employees[employee.id].department;
					if (!func.setNotNull(department)) continue;
					if (!department.toLowerCase().includes(settings.searchQuery.toLowerCase())) continue;
				}
				else if (settings.searchType == 'By Location') {
					let office = settings.employees[employee.id].office || stored.employees[employee.id].office;

					if (!func.setNotNull(office)) continue;
					if (!office.toLowerCase().includes(settings.searchQuery.toLowerCase())) continue;
				}
				else if (settings.searchType == 'By Job Title') {
					let jobTitle = settings.employees[employee.id].jobTitle || stored.employees[employee.id].jobTitle;

					if (!func.setNotNull(jobTitle)) continue;
					if (!jobTitle.toLowerCase().includes(settings.searchQuery.toLowerCase())) continue;
				}
			}

			let photo = this.sharePoint.images.user;
			let image = settings.employees[employee.id].photo;
			if (func.setNotNull(settings.employees[employee.id].photo)) {
				photo = window.URL.createObjectURL(image);
			}

			display.makeElement({
				element: 'div', attributes: { class: 'crater-employee-directory-row', id: `row-${employee.id}` }, children: [
					{ element: 'img', attributes: { class: 'crater-employee-directory-dp', src: photo } },
					{
						element: 'span', attributes: { class: 'crater-employee-directory-details' }, children: [
							{ element: 'p', attributes: { class: 'crater-employee-directory-name' }, text: employee.displayName },
							{ element: 'p', attributes: { class: 'crater-employee-directory-mail' }, text: employee.mail },
							{ element: 'p', attributes: { class: 'crater-employee-directory-job' }, text: employee.jobTitle },
							{ element: 'p', attributes: { class: 'crater-employee-directory-contact' } },
							{
								element: 'span', attributes: { class: 'crater-employee-directory-contact' }, children: [
									{ element: 'img', attributes: { class: 'crater-employee-directory-icon crater-employee-directory-mail', src: 'https://banner2.cleanpng.com/20180720/ixe/kisspng-computer-icons-email-icon-design-equipo-comercial-5b525b3cdb7d21.311695661532123964899.jpg' } },
									{ element: 'img', attributes: { class: 'crater-employee-directory-icon crater-employee-directory-message', src: 'https://www.pinclipart.com/picdir/middle/107-1070124_message-png-clipart-computer-icons-clip-art-transparent.png' } },
									{ element: 'img', attributes: { class: 'crater-employee-directory-icon crater-employee-directory-phone', src: 'https://p7.hiclipart.com/preview/211/783/729/telephone-symbol-icon-phone-download-png.jpg' } }
								]
							}
						]
					},
					{ element: 'img', attributes: { class: 'crater-employee-directory-icon crater-employee-directory-toggle-view', src: this.openImage } }
				]
			});
		}

		localStorage[`crater-${this.key}`] = JSON.stringify(settings);
	}

	public setUpPaneContent(params): any {
		this.element = params.element;
		this.key = params.element.dataset['key'];

		this.paneContent = this.elementModifier.createElement({
			element: 'div',
			attributes: { class: 'crater-property-content' }
		});

		if (this.sharePoint.properties.pane.content[this.key].draft.pane.content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.pane.content;
		}
		else if (this.sharePoint.properties.pane.content[this.key].content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].content;
		}
		else {
			let menuPane = this.paneContent.makeElement({
				element: 'div', attributes: { class: 'card menu-pane' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: "Menu Settings"
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'Background Color', list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Border Size', list: func.pixelSizes
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Border Color', list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Border Type', list: func.borderTypes
							})
						]
					})
				]
			});

			let searchTypePane = this.paneContent.makeElement({
				element: 'div', attributes: { class: 'card search-type-pane' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: "Search Type Settings"
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'Shadow', list: func.shadows
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Border', list: func.borders
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Color', list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Background Color', list: func.colors
							})
						]
					})
				]
			});

			let displayPane = this.paneContent.makeElement({
				element: 'div', attributes: { class: 'card display-pane' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: "Search Result Settings"
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'Height', list: func.pixelSizes
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Background Color', list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Font Color', list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Font Size', list: func.pixelSizes
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Font Style', list: func.fontStyles
							}),
							this.elementModifier.cell({
								element: 'img', name: 'Default Avatar'
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Avater Background', list: func.colors
							})
						]
					})
				]
			});

			let appsPane = this.paneContent.makeElement({
				element: 'div', attributes: { class: 'card apps-pane' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: "Default Apps"
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'select', name: 'Mail', options: ['Outlook', 'Gmail', 'Yahoo']
							}),
							this.elementModifier.cell({
								element: 'select', name: 'Message', options: ['Teams', 'Skype']
							}),
							this.elementModifier.cell({
								element: 'select', name: 'Call', options: ['Teams', 'Skype']
							}),
						]
					})
				]
			});
		}

		return this.paneContent;
	}

	public listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset.key;
		this.paneContent = this.sharePoint.app.find('.crater-property-content').monitor();
		let settings = this.sharePoint.properties.pane.content[this.key].settings;

		let domDraft = this.sharePoint.properties.pane.content[this.key].draft.dom;

		let menuPane = this.paneContent.find('.menu-pane');
		let searchTypePane = this.paneContent.find('.search-type-pane');
		let displayPane = this.paneContent.find('.display-pane');
		let appsPane = this.paneContent.find('.apps-pane');

		appsPane.find('#Mail-cell').onChanged();
		appsPane.find('#Message-cell').onChanged();
		appsPane.find('#Call-cell').onChanged();

		menuPane.find('#Border-Color-cell').onChanged(borderColor => {
			let borderType = menuPane.find('#Border-Type-cell').value || 'solid';
			let borderSize = menuPane.find('#Border-Size-cell').value || '1px';
			domDraft.find('.crater-employee-directory-menu').css({ border: `${borderSize} ${borderType} ${borderColor}` });
		});

		menuPane.find('#Border-Size-cell').onChanged(borderSize => {
			let borderType = menuPane.find('#Border-Type-cell').value || 'solid';
			let borderColor = menuPane.find('#Border-Color-cell').value || '1px';
			domDraft.find('.crater-employee-directory-menu').css({ border: `${borderSize} ${borderType} ${borderColor}` });
		});

		menuPane.find('#Border-Type-cell').onChanged(borderType => {
			let borderSize = menuPane.find('#Border-Size-cell').value || 'solid';
			let borderColor = menuPane.find('#Border-Color-cell').value || '1px';
			domDraft.find('.crater-employee-directory-menu').css({ border: `${borderSize} ${borderType} ${borderColor}` });
		});

		menuPane.find('#Background-Color-cell').onChanged(backgroundColor => {
			domDraft.find('.crater-employee-directory-menu').css({ backgroundColor });
		});

		searchTypePane.find('#Shadow-cell').onChanged(boxShadow => {
			domDraft.find('#crater-employee-directory-search-type').css({ boxShadow });
		});

		searchTypePane.find('#Border-cell').onChanged(border => {
			domDraft.find('#crater-employee-directory-search-type').css({ border });
		});

		searchTypePane.find('#Color-cell').onChanged(color => {
			domDraft.find('#crater-employee-directory-search-type').css({ color });
		});

		searchTypePane.find('#Background-Color-cell').onChanged(backgroundColor => {
			domDraft.find('#crater-employee-directory-search-type').css({ backgroundColor });
		});

		displayPane.find('#Height-cell').onChanged(height => {
			domDraft.find('.crater-employee-directory-display').css({ height });
		});

		displayPane.find('#Background-Color-cell').onChanged(backgroundColor => {
			domDraft.find('.crater-employee-directory-display').css({ backgroundColor });
		});

		displayPane.find('#Font-Color-cell').onChanged(color => {
			domDraft.find('.crater-employee-directory-display').css({ color });
		});

		displayPane.find('#Font-Size-cell').onChanged(fontSize => {
			domDraft.find('.crater-employee-directory-display').css({ fontSize });
		});

		displayPane.find('#Font-Style-cell').onChanged(fontFamily => {
			domDraft.find('.crater-employee-directory-display').css({ fontFamily });
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());
			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart

			this.sharePoint.properties.pane.content[this.key].settings.mailApp = appsPane.find('#Mail-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.messageApp = appsPane.find('#Message-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.callApp = appsPane.find('#Call-cell').value;
		});
	}
}

class Carousel extends CraterWebParts {
	public params: any;
	public elementModifier = new ElementModifier();
	public paneContent: any;
	public element: any;
	private backgroundImage = '';
	private backgroundColor = '#999999';
	private color = 'white';
	private columns: number = 3;
	private duration: number = 10;
	private key: any;

	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.backgroundImage = this.sharePoint.images.append;
		this.params = params;
	}

	public render(params) {
		if (!func.isset(params.source)) params.source = [
			{ image: this.sharePoint.images.append, text: 'One here' },
			{ image: this.sharePoint.images.edit, text: 'Two now' },
			{ image: this.sharePoint.images.sync, text: 'Three then' },
			{ image: this.sharePoint.images.async, text: 'Four done' },
			{ image: this.sharePoint.images.delete, text: 'Five when' }
		];

		let carousel = this.createKeyedElement({ element: 'div', attributes: { class: 'crater-carousel crater-component', 'data-type': 'carousel' } });

		let radioToggle = carousel.makeElement({ element: 'span', attributes: { class: 'crater-top-right', id: 'crater-carousel-controller' } });

		let content = carousel.makeElement({
			element: 'div', attributes: { class: 'crater-carousel-content' }
		});

		for (let src of params.source) {
			radioToggle.makeElement({ element: 'input', attributes: { type: 'radio', class: 'crater-carousel-radio-toggle' } });
			content.makeElement({
				element: 'span', attributes: { class: 'crater-carousel-column' }, children: [
					{ element: 'img', attributes: { class: 'crater-carousel-image', src: src.image } },
					{ element: 'span', attributes: { class: 'crater-carousel-text' }, text: src.text }
				]
			});
		}

		carousel.makeElement({ element: 'span', attributes: { class: 'crater-arrow crater-left-arrow' } });
		carousel.makeElement({ element: 'span', attributes: { class: 'crater-arrow crater-right-arrow' } });

		this.key = this.key || carousel.dataset.key;
		//upload the pre-defined settings of the webpart
		this.sharePoint.properties.pane.content[this.key].settings.columns = 3;
		this.sharePoint.properties.pane.content[this.key].settings.duration = 1000;
		return carousel;
	}

	public rendered(params) {
		this.params = params;
		this.element = params.element;
		this.key = this.element.dataset['key'];

		this.columns = this.sharePoint.properties.pane.content[this.key].settings.columns / 1;

		this.duration = this.sharePoint.properties.pane.content[this.key].settings.duration / 1;

		this.startSlide();
	}

	public startSlide() {
		this.key = this.element.dataset['key'];
		let controller = this.element.find('#crater-carousel-controller'),
			arrows = this.element.findAll('.crater-arrow'),
			radios,
			columns = this.element.findAll('.crater-carousel-column'),
			radio: any,
			key = 0;

		if (this.element.length < 1) return;

		//reset control buttons
		controller.innerHTML = '';

		//stack the first slide ontop
		for (let position = 0; position < columns.length; position++) {
			columns[position].css({ zIndex: 0 });
			if (position == 0) columns[position].css({ zIndex: 1 });
			controller.makeElement({
				element: 'input', attributes: { class: 'crater-carousel-radio-toggle', type: 'radio' }
			});
		}
		radios = controller.findAll('.crater-carousel-radio-toggle');

		//fading and fadeout animation

		let runSliding = () => {
			if (key < 0) key = radios.length - 1;
			if (key >= radios.length) key = 0;
			for (let element of radios) {
				if (radio != element) element.checked = false;
			}

			let getColumns = () => {
				let currentColunms = [];
				for (let i = 0; i < this.columns; i++) {
					currentColunms[i] = key + i;
					if (currentColunms[i] >= columns.length) {
						currentColunms[i] -= columns.length;
					}
				}

				return currentColunms;
			};

			let current = getColumns();
			for (let i = 0; i < columns.length; i++) {
				let position: number = current.indexOf(i);
				if (position != -1) {
					columns[i].show();
					columns[i].css({ gridColumnStart: position + 1, gridRowStart: 1 });
				} else {
					columns[i].hide();
				}
			}
		};

		//move to next slide
		let keepSliding = () => {
			clearInterval(this.sharePoint.properties.pane.content[this.key].settings.animation);
			this.sharePoint.properties.pane.content[this.key].settings.animation = setInterval(() => {
				key++;
				runSliding();
			}, this.sharePoint.properties.pane.content[this.key].settings.duration);
		};

		//run animation when arrow is clicked
		for (let arrow of arrows) {
			arrow.css({ marginTop: `${(this.element.position().height - arrow.position().height) / 2}px` });
			arrow.addEventListener('click', event => {
				if (arrow.classList.contains('crater-left-arrow')) {
					key--;
				}
				else if (arrow.classList.contains('crater-right-arrow')) {
					key++;
				}

				clearInterval(this.sharePoint.properties.pane.content[this.key].settings.animation);
				runSliding();
			});
		}

		//run animation when a controller is clicked
		for (let position = 0; position < radios.length; position++) {
			radios[position].addEventListener('click', () => {
				clearInterval(this.sharePoint.properties.pane.content[this.key].settings.animation);
				key = position;
				runSliding();
			});
		}

		//click the first controller and set the first slide
		radios[key].click();
		keepSliding();
	}

	public setUpPaneContent(params): any {
		this.element = params.element;
		this.key = params.element.dataset['key'];

		this.paneContent = this.elementModifier.createElement({
			element: 'div',
			attributes: { class: 'crater-property-content' }
		});

		if (this.sharePoint.properties.pane.content[this.key].draft.pane.content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.pane.content;
		}
		else if (this.sharePoint.properties.pane.content[this.key].content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].content;
		}
		else {
			let columns = this.sharePoint.properties.pane.content[this.key].draft.dom.findAll('.crater-carousel-column');

			this.paneContent.makeElement({
				element: 'div', children: [
					this.elementModifier.createElement({
						element: 'button', attributes: { class: 'btn new-component', style: { display: 'inline-block', borderRadius: '5px' } }, text: 'Add New'
					})
				]
			});

			this.paneContent.append(this.generatePaneContent({ columns }));

			let settingsPane = this.paneContent.makeElement({
				element: 'div', attributes: { class: 'card settings-pane' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: "Settings"
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'Duration', value: this.sharePoint.properties.pane.content[this.key].settings.duration || ''
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Columns', value: this.sharePoint.properties.pane.content[this.key].settings.columns || ''
							}),
							this.elementModifier.cell({
								element: 'input', name: 'FontSize'
							}),
							this.elementModifier.cell({
								element: 'input', name: 'FontStyle'
							}),
							this.elementModifier.cell({
								element: 'input', name: 'ImageSize'
							}),
							this.elementModifier.cell({
								element: 'select', name: 'ShowText', value: this.sharePoint.properties.pane.content[this.key].settings.showText || '', options: ['Yes', 'No']
							}),
							this.elementModifier.cell({
								element: 'select', name: 'Curved', value: this.sharePoint.properties.pane.content[this.key].settings.curved || '', options: ['Yes', 'No']
							}),
							this.elementModifier.cell({
								element: 'select', name: 'Shadow', value: this.sharePoint.properties.pane.content[this.key].settings.shadow || '', options: ['Yes', 'No']
							}),
						]
					})
				]
			});
		}

		return this.paneContent;
	}

	public generatePaneContent(params) {
		let columnsPane = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'card columns-pane' }, children: [
				this.elementModifier.createElement({
					element: 'div', attributes: { class: 'card-title' }, children: [
						this.elementModifier.createElement({
							element: 'h2', attributes: { class: 'title' }, text: "Columns"
						})
					]
				}),
			]
		});

		for (let i = 0; i < params.columns.length; i++) {
			columnsPane.makeElement({
				element: 'div',
				attributes: {
					style: { border: '1px solid #bbbbbb', margin: '.5em 0em' }, class: 'crater-carousel-column-pane row'
				},
				children: [
					this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-carousel-column' }),
					this.elementModifier.cell({
						element: 'img', name: 'Image', attributes: {}, dataAttributes: { class: 'crater-icon', src: params.columns[i].find('.crater-carousel-image').src }
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Text', attributes: {}, value: params.columns[i].find('.crater-carousel-text').innerText || ''
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Color', attributes: {}, value: params.columns[i].find('.crater-carousel-text').css().color || '', list: func.colors
					}),
					this.elementModifier.cell({
						element: 'input', name: 'BackgroundColor', attributes: {}, value: params.columns[i].find('.crater-carousel-text').css().backgroundColor || ''
					}),
				]
			});
		}

		return columnsPane;
	}

	public listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		this.paneContent = this.sharePoint.app.find('.crater-property-content').monitor();

		let domDraft = this.sharePoint.properties.pane.content[this.key].draft.dom;
		let content = domDraft.find('.crater-carousel-content');
		let columns = domDraft.findAll('.crater-carousel-column');

		let columnPanePrototype = this.elementModifier.createElement({
			element: 'div',
			attributes: {
				style: { border: '1px solid #bbbbbb', margin: '.5em 0em' }, class: 'crater-carousel-column-pane row'
			},
			children: [
				this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-carousel-column' }),
				this.elementModifier.cell({
					element: 'img', name: 'Image', attributes: {}, dataAttributes: { class: 'crater-icon', src: this.sharePoint.images.append }
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Text', attributes: {}, value: 'Text Here'
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Color', attributes: {}, list: func.colors
				}),
				this.elementModifier.cell({
					element: 'input', name: 'BackgroundColor', attributes: {}, list: func.colors
				}),
			]
		});

		let columnPrototype = this.elementModifier.createElement({
			element: 'span', attributes: { class: 'crater-carousel-column' }, children: [
				{ element: 'img', attributes: { class: 'crater-carousel-image', src: this.sharePoint.images.append } },
				{ element: 'span', attributes: { class: 'crater-carousel-text' }, text: 'Text Here' }
			]
		});

		let carouselColumnHandler = (columnPane, columnDom) => {
			columnPane.addEventListener('mouseover', event => {
				columnPane.find('.crater-content-options').css({ visibility: 'visible' });
			});

			columnPane.addEventListener('mouseout', event => {
				columnPane.find('.crater-content-options').css({ visibility: 'hidden' });
			});

			let imageCell = columnPane.find('#Image-cell').parentNode;
			this.uploadImage({ parent: imageCell }, (image) => {
				imageCell.find('#Image-cell').src = image.src;
				columnDom.find('.crater-carousel-image').src = image.src;
			});

			columnPane.find('#Text-cell').onChanged(value => {
				columnDom.find('.crater-carousel-text').textContent = value;
			});

			let colorCell = columnPane.find('#Color-cell').parentNode;
			this.pickColor({ parent: colorCell, cell: colorCell.find('#Color-cell') }, (color) => {
				columnDom.find('.crater-carousel-text').css({ color });
				colorCell.find('#Color-cell').value = color;
				colorCell.find('#Color-cell').setAttribute('value', color);
			});

			let backgroundColorCell = columnPane.find('#BackgroundColor-cell').parentNode;
			this.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.find('#BackgroundColor-cell') }, (backgroundColor) => {
				columnDom.css({ backgroundColor });
				backgroundColorCell.find('#BackgroundColor-cell').value = backgroundColor;
				backgroundColorCell.find('#BackgroundColor-cell').setAttribute('value', backgroundColor);
			});

			columnPane.find('.delete-crater-carousel-column').addEventListener('click', event => {
				columnDom.remove();
				columnPane.remove();
			});

			columnPane.find('.add-before-crater-carousel-column').addEventListener('click', event => {
				let newColumnPrototype = columnPrototype.cloneNode(true);
				let newColumnPanePrototype = columnPanePrototype.cloneNode(true);

				columnDom.before(newColumnPrototype);
				columnPane.before(newColumnPanePrototype);
				carouselColumnHandler(newColumnPanePrototype, newColumnPrototype);
			});

			columnPane.find('.add-after-crater-carousel-column').addEventListener('click', event => {
				let newColumnPrototype = columnPrototype.cloneNode(true);
				let newColumnPanePrototype = columnPanePrototype.cloneNode(true);

				columnDom.after(newColumnPrototype);
				columnPane.after(newColumnPanePrototype);
				carouselColumnHandler(newColumnPanePrototype, newColumnPrototype);
			});
		};

		this.paneContent.find('.new-component').addEventListener('click', event => {
			let newColumnPrototype = columnPrototype.cloneNode(true);
			let newColumnPanePrototype = columnPanePrototype.cloneNode(true);

			content.append(newColumnPrototype);//c
			this.paneContent.find('.columns-pane').append(newColumnPanePrototype);

			carouselColumnHandler(newColumnPanePrototype, newColumnPrototype);
		});

		this.paneContent.findAll('.crater-carousel-column-pane').forEach((columnPane, position) => {
			carouselColumnHandler(columnPane, columns[position]);
		});

		let settingsPane = this.paneContent.find('.settings-pane');

		settingsPane.find('#Duration-cell').onChanged();

		settingsPane.find('#Columns-cell').onChanged(value => {
			domDraft.find('.crater-carousel-content').css({ gridTemplateColumns: `repeat(${value}, 1fr)` });
		});

		settingsPane.find('#FontSize-cell').onChanged(fontSize => {
			domDraft.findAll('.crater-carousel-text').forEach(text => {
				text.css({ fontSize });
			});
		});

		settingsPane.find('#FontStyle-cell').onChanged(fontFamily => {
			domDraft.findAll('.crater-carousel-text').forEach(text => {
				text.css({ fontFamily });
			});
		});

		settingsPane.find('#ImageSize-cell').onChanged(width => {
			domDraft.findAll('.crater-carousel-image').forEach(text => {
				text.css({ width });
			});
		});

		settingsPane.find('#ShowText-cell').onChanged(display => {
			domDraft.findAll('.crater-carousel-text').forEach(text => {
				if (display.toLowerCase() == 'no') {
					text.hide();
				} else {
					text.show();
				}
			});
		});

		settingsPane.find('#Curved-cell').onChanged(curved => {
			domDraft.findAll('.crater-carousel-column').forEach(column => {
				if (curved.toLowerCase() == 'yes') {
					column.css({ borderRadius: '10px' });
				} else {
					column.cssRemove(['border-radius']);
				}
			});
		});

		settingsPane.find('#Shadow-cell').onChanged(shadow => {
			domDraft.findAll('.crater-carousel-column').forEach(column => {
				if (shadow.toLowerCase() == 'yes') {
					column.css({ boxShadow: 'var(--accient-shadow)' });
				} else {
					column.cssRemove(['box-shadow']);
				}
			});
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());
			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart

			this.sharePoint.properties.pane.content[this.key].settings.duration = this.paneContent.find('#Duration-cell').value;
			this.sharePoint.properties.pane.content[this.key].settings.columns = this.paneContent.find('#Columns-cell').value;
		});
	}

	public update(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;
		this.paneContent = this.setUpPaneContent(params);

		let paneConnection = this.sharePoint.app.find('.crater-property-connection');

		let updateWindow = this.elementModifier.createForm({
			title: 'Setup Meta Data', attributes: { id: 'meta-data-form', class: 'form' },
			contents: {
				image: { element: 'select', attributes: { id: 'meta-data-image', name: 'Image' }, options: params.options },
				text: { element: 'select', attributes: { id: 'meta-data-text', name: 'Text' }, options: params.options },
			},
			buttons: {
				submit: { element: 'button', attributes: { id: 'update-element', class: 'btn' }, text: 'Update' },
			}
		});

		let data: any = {};
		let source: any;
		updateWindow.find('#update-element').addEventListener('click', event => {
			event.preventDefault();
			data.image = updateWindow.find('#meta-data-image').value;
			data.text = updateWindow.find('#meta-data-text').value;
			source = func.extractFromJsonArray(data, params.source);

			let newContent = this.render({ source });
			draftDom.find('.crater-carousel-content').innerHTML = newContent.find('.crater-carousel-content').innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = draftDom.outerHTML;

			this.paneContent.find('.columns-pane').innerHTML = this.generatePaneContent({ columns: newContent.findAll('.crater-carousel-column') }).innerHTML;

			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
		});


		if (!func.isnull(paneConnection)) {
			paneConnection.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
				this.element.innerHTML = draftDom.innerHTML;

				this.element.css(draftDom.css());

				this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart
			});
		}

		return updateWindow;
	}
}

class Events extends CraterWebParts {
	public params: any;
	public elementModifier = new ElementModifier();
	public paneContent: any;
	public element: any;
	private key: any;

	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.params = params;
	}

	private render(params) {
		if (!func.isset(params.source)) params.source = [
			{ task: 'Facebook', date: func.today(), icon: this.sharePoint.images.append, location: 'London', duration: func.time() },
			{ task: 'Facebook', date: func.today(), icon: this.sharePoint.images.append, location: 'London', duration: func.time() },
			{ task: 'Facebook', date: func.today(), icon: this.sharePoint.images.append, location: 'London', duration: func.time() },
		];

		let events = this.createKeyedElement({
			element: 'div', attributes: { class: 'crater-events crater-components', 'data-type': 'events' }, children: [
				{
					element: 'div', attributes: { class: 'crater-events-title' }, children: [
						{ element: 'img', attributes: { class: 'crater-events-title-icon', src: this.sharePoint.images.append } },
						{ element: 'span', attributes: { class: 'crater-events-title-text' }, text: 'Events' }
					]
				},
				{ element: 'div', attributes: { class: 'crater-events-content' } }
			]
		});

		let content = events.find('.crater-events-content');

		for (let row of params.source) {
			content.makeElement({
				element: 'div', attributes: { class: 'crater-events-row' }, children: [
					{ element: 'img', attributes: { class: 'crater-events-row-icon', src: row.icon, id: 'icon' } },
					{
						element: 'span', attributes: { class: 'crater-events-details' }, children: [
							{ element: 'span', attributes: { class: 'crater-events-task', id: 'task' }, text: row.task },
							{
								element: 'span', attributes: { class: 'crater-events-durationandlocation' }, children: [
									{
										element: 'span', attributes: { class: 'crater-events-duration' }, children: [
											{ element: 'img', attributes: { class: 'crater-events-duration-icon', src: this.sharePoint.images.duration } },
											{ element: 'h5', attributes: { class: 'crater-events-duration-value', id: 'duration' }, text: row.duration }
										]
									},
									{
										element: 'span', attributes: { class: 'crater-events-location' }, children: [
											{ element: 'img', attributes: { class: 'crater-events-location-icon', src: this.sharePoint.images.location } },
											{ element: 'h5', attributes: { class: 'crater-events-location-value', id: 'location' }, text: row.location }
										]
									}
								]
							}
						]
					},
					{
						element: 'span', attributes: { class: 'crater-events-date' }, children: [
							{ element: 'a', attributes: { class: 'crater-events-date-value', id: 'date' }, text: row.date }
						]
					}
				]
			});
		}

		content.css({ gridTemplateRows: `repeat${params.source.length}, 1fr` });

		this.key = this.key || events.dataset.key;
		this.sharePoint.properties.pane.content[this.key].settings.layout = 'Icon, Details, Date';
		return events;
	}

	private rendered(params) {
		this.element = params.element;
		this.key = params.element.dataset.key;

		let layout = this.sharePoint.properties.pane.content[this.key].settings.layout.split(', ');
		let iconPosition = layout.indexOf('Icon') + 1;
		let detailsPosition = layout.indexOf('Details') + 1;
		let datePosition = layout.indexOf('Date') + 1;

		let rows = this.element.findAll('.crater-events-row');
		rows.forEach(row => {
			row.find('.crater-events-row-icon').css({ gridColumnStart: iconPosition, gridColumnEnd: iconPosition, gridRowStart: 1 });
			row.find('.crater-events-details').css({ gridColumnStart: detailsPosition, gridColumnEnd: detailsPosition, gridRowStart: 1 });
			row.find('.crater-events-date').css({ gridColumnStart: datePosition, gridColumnEnd: datePosition, gridRowStart: 1 });
		});

	}

	public setUpPaneContent(params): any {
		let key = params.element.dataset['key'];
		this.element = this.sharePoint.properties.pane.content[key].draft.dom;
		this.paneContent = this.elementModifier.createElement({
			element: 'div',
			attributes: { class: 'crater-property-content' }
		}).monitor();

		if (this.sharePoint.properties.pane.content[key].draft.pane.content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[key].draft.pane.content;
		}
		else if (this.sharePoint.properties.pane.content[key].content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[key].content;
		}
		else {
			let eventsContent = this.sharePoint.properties.pane.content[key].draft.dom.find('.crater-events-content');
			let eventsRows = eventsContent.findAll('.crater-events-row');

			this.paneContent.makeElement({
				element: 'div', children: [
					this.elementModifier.createElement({
						element: 'button', attributes: { class: 'btn new-component', style: { display: 'inline-block', borderRadius: '5px' } }, text: 'Add New'
					})
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'title-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Events Title'
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'img', name: 'icon', attributes: {}, dataAttributes: { class: 'crater-icon', src: this.element.find('.crater-events-title-icon').src }
							}),
							this.elementModifier.cell({
								element: 'input', name: 'iconsize', value: this.element.find('.crater-events-title-icon').css()['width'] || ''
							}),
							this.elementModifier.cell({
								element: 'input', name: 'title', value: this.element.find('.crater-events-title-text').textContent
							}),
							this.elementModifier.cell({
								element: 'input', name: 'backgroundcolor', value: this.element.find('.crater-events-title').css()['background-color'], list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'color', value: this.element.find('.crater-events-title').css().color, list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontsize', value: this.element.find('.crater-events-title').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontstyle', value: this.element.find('.crater-events-title').css()['font-family'] || ''
							}),
							this.elementModifier.cell({
								element: 'input', name: 'height', value: this.element.find('.crater-events-title').css()['height'] || ''
							})
						]
					})
				]
			});

			this.paneContent.append(this.generatePaneContent({ events: eventsRows }));

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'icons-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Events Icons'
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'size'
							}),
							this.elementModifier.cell({
								element: 'select', name: 'show', options: ['Yes', 'No']
							}),
						]
					})
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'tasks-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Events Tasks'
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'fontsize'
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontstyle'
							}),
							this.elementModifier.cell({
								element: 'input', name: 'color', list: func.colors
							}),
							this.elementModifier.cell({
								element: 'select', name: 'show', options: ['Yes', 'No']
							}),
						]
					})
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'durations-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Events Durations'
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'fontsize'
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontstyle'
							}),
							this.elementModifier.cell({
								element: 'input', name: 'color'
							}),
							this.elementModifier.cell({
								element: 'select', name: 'show', options: ['Yes', 'No']
							}),
						]
					})
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'locations-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Events Locations'
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'fontsize'
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontstyle'
							}),
							this.elementModifier.cell({
								element: 'input', name: 'color'
							}),
							this.elementModifier.cell({
								element: 'select', name: 'show', options: ['Yes', 'No']
							}),
						]
					})
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'dates-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Events Dates'
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'fontsize'
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontstyle'
							}),
							this.elementModifier.cell({
								element: 'input', name: 'color'
							}),
							this.elementModifier.cell({
								element: 'select', name: 'show', options: ['Yes', 'No']
							}),
						]
					})
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'settings-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Events Settings'
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'select', name: 'Layout', options: ['Icon, Details, Date', 'Icon, Date, Details', 'Date, Details, Icon', 'Date, Icon, Details', 'Details, Date, Icon', 'Details, Icon, Date']
							})
						]
					})
				]
			});
		}

		return this.paneContent;
	}

	public generatePaneContent(params) {
		let eventsPane = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'card events-pane' }, children: [
				this.elementModifier.createElement({
					element: 'div', attributes: { class: 'card-title' }, children: [
						this.elementModifier.createElement({
							element: 'h2', attributes: { class: 'title' }, text: 'Events Content'
						})
					]
				}),
			]
		});

		for (let i = 0; i < params.events.length; i++) {
			eventsPane.makeElement({
				element: 'div',
				attributes: { class: 'crater-events-row-pane row' },
				children: [
					this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-event-row' }),
					this.elementModifier.cell({
						element: 'img', name: 'icon', attributes: {}, dataAttributes: { class: 'crater-icon', src: params.events[i].find('#icon').src }
					}),
					this.elementModifier.cell({
						element: 'input', name: 'task', value: params.events[i].find('#task').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'duration', value: params.events[i].find('#duration').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'location', value: params.events[i].find('#location').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'date', value: params.events[i].find('#date').textContent
					}),
				]
			});
		}

		return eventsPane;
	}

	private listenPaneContent(params) {
		this.key = params.element.dataset['key'];
		this.element = params.element;
		let domDraft = this.sharePoint.properties.pane.content[this.key].draft.dom;
		this.paneContent = this.sharePoint.app.find('.crater-property-content').monitor();

		let eventPrototype = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'crater-events-row' }, children: [
				{ element: 'img', attributes: { class: 'crater-events-row-icon', src: this.sharePoint.images.append, id: 'icon' } },
				{
					element: 'span', attributes: { class: 'crater-events-details' }, children: [
						{ element: 'span', attributes: { class: 'crater-events-task', id: 'task' }, text: "Task" },
						{
							element: 'span', attributes: { class: 'crater-events-durationandlocation' }, children: [
								{
									element: 'span', attributes: { class: 'crater-events-duration' }, children: [
										{ element: 'img', attributes: { class: 'crater-events-duration-icon', src: this.sharePoint.images.duration } },
										{ element: 'h5', attributes: { class: 'crater-events-duration-value', id: 'duration' }, text: "1st December - 4th December" }
									]
								},
								{
									element: 'span', attributes: { class: 'crater-events-location' }, children: [
										{ element: 'img', attributes: { class: 'crater-events-location-icon', src: this.sharePoint.images.location } },
										{ element: 'h5', attributes: { class: 'crater-events-location-value', id: 'location' }, text: "Lagos" }
									]
								}
							]
						}
					]
				},
				{
					element: 'span', attributes: { class: 'crater-events-date' }, children: [
						{ element: 'a', attributes: { class: 'crater-events-date-value', id: 'date' }, text: func.today() }
					]
				}
			]
		});

		let eventPanePrototype = this.elementModifier.createElement({
			element: 'div',
			attributes: { class: 'crater-events-row-pane row' },
			children: [
				this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-event-row' }),
				this.elementModifier.cell({
					element: 'img', name: 'icon', attributes: {}, dataAttributes: { class: 'crater-icon', src: this.sharePoint.images.append }
				}),
				this.elementModifier.cell({
					element: 'input', name: 'task', value: 'Task'
				}),
				this.elementModifier.cell({
					element: 'input', name: 'duration', value: '1st December - 4th December'
				}),
				this.elementModifier.cell({
					element: 'input', name: 'location', value: 'Lagos'
				}),
				this.elementModifier.cell({
					element: 'input', name: 'date', value: func.today()
				}),
			]
		});

		let titlePane = this.paneContent.find('.title-pane');

		titlePane.find('#title-cell').onChanged(value => {
			domDraft.find('.crater-events-title-text').textContent = value;
		});

		titlePane.find('#fontsize-cell').onChanged(fontSize => {
			domDraft.find('.crater-events-title-text').css({ fontSize });
		});

		titlePane.find('#fontstyle-cell').onChanged(fontFamily => {
			domDraft.find('.crater-events-title-text').css({ fontFamily });
		});

		let colorCell = titlePane.find('#color-cell').parentNode;
		this.pickColor({ parent: colorCell, cell: colorCell.find('#color-cell') }, (color) => {
			domDraft.find('.crater-events-title-text').css({ color });
			colorCell.find('#color-cell').value = color;
			colorCell.find('#color-cell').setAttribute('value', color);
		});

		let iconCell = titlePane.find('#icon-cell').parentNode;
		this.uploadImage({ parent: iconCell }, (image) => {
			iconCell.find('#icon-cell').src = image.src;
			domDraft.find('.crater-events-title-icon').src = image.src;
		});

		titlePane.find('#iconsize-cell').onChanged(width => {
			domDraft.find('.crater-events-title-icon').css({ width });
		});

		titlePane.find('#height-cell').onChanged(height => {
			domDraft.find('.crater-events-title').css({ height });
		});

		let backgroundColorCell = titlePane.find('#backgroundcolor-cell').parentNode;
		this.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.find('#backgroundcolor-cell') }, (backgroundColor) => {
			domDraft.find('.crater-events-title').css({ backgroundColor });
			backgroundColorCell.find('#backgroundcolor-cell').value = backgroundColor;
			backgroundColorCell.find('#backgroundcolor-cell').setAttribute('value', backgroundColor);
		});

		let eventHandler = (eventPane, eventDom) => {
			eventPane.addEventListener('mouseover', event => {
				eventPane.find('.crater-content-options').css({ visibility: 'visible' });
			});

			eventPane.addEventListener('mouseout', event => {
				eventPane.find('.crater-content-options').css({ visibility: 'hidden' });
			});

			eventPane.find('.delete-crater-event-row').addEventListener('click', event => {
				eventDom.remove();
				eventPane.remove();
			});

			eventPane.find('.add-before-crater-event-row').addEventListener('click', event => {
				let newEventventPrototype = eventPrototype.cloneNode(true);
				let newEventPanePrototype = eventPanePrototype.cloneNode(true);

				eventDom.before(newEventventPrototype);
				eventPane.before(newEventPanePrototype);
				eventHandler(newEventPanePrototype, newEventventPrototype);
			});

			eventPane.find('.add-after-crater-event-row').addEventListener('click', event => {
				let newEventventPrototype = eventPrototype.cloneNode(true);
				let newEventPanePrototype = eventPanePrototype.cloneNode(true);

				eventDom.after(newEventventPrototype);
				eventPane.after(newEventPanePrototype);
				eventHandler(newEventPanePrototype, newEventventPrototype);
			});

			let eventIconCell = eventPane.find('#icon-cell').parentNode;
			this.uploadImage({ parent: eventIconCell }, (image) => {
				eventIconCell.find('#icon-cell').src = image.src;
				domDraft.find('.crater-events-row-icon').src = image.src;
			});

			eventPane.find('#task-cell').onChanged(value => {
				eventDom.find('.crater-events-task').innerText = value;
			});

			eventPane.find('#duration-cell').onChanged(value => {
				eventDom.find('.crater-events-duration').innerText = value;
			});

			eventPane.find('#location-cell').onChanged(value => {
				eventDom.find('.crater-events-location').innerText = value;
			});
			eventPane.find('#date-cell').onChanged(value => {
				eventDom.find('.crater-events-date').innerText = value;
			});
		};

		let eventsPane = this.paneContent.find('.events-pane');

		this.paneContent.find('.new-component').addEventListener('click', event => {
			let newevEntventPrototype = eventPrototype.cloneNode(true);
			let newEventPanePrototype = eventPanePrototype.cloneNode(true);

			eventsPane.append(newEventPanePrototype);
			domDraft.find('.crater-events-content').append(newevEntventPrototype);
			eventHandler(newEventPanePrototype, newevEntventPrototype);
		});

		eventsPane.findAll('.crater-events-row-pane').forEach((eventPane, position) => {
			eventHandler(eventPane, domDraft.findAll('.crater-events-row')[position]);
		});

		let iconsPane = this.paneContent.find('.icons-pane');

		iconsPane.find('#size-cell').onChanged(width => {
			domDraft.findAll('.crater-events-row-icon').forEach(icon => {
				icon.css({ width });
			});
		});

		iconsPane.find('#show-cell').onChanged(display => {
			domDraft.findAll('.crater-events-row-icon').forEach(icon => {
				if (display.toLowerCase() == 'no') icon.css({ display: 'none' });
				else icon.cssRemove(['display']);
			});
		});

		let handleProperties = (property, pane) => {
			pane.find('#fontsize-cell').onChanged(fontSize => {
				domDraft.findAll(`.crater-events-${property}-value`).forEach(value => {
					value.css({ fontSize });
				});
			});

			pane.find('#fontstyle-cell').onChanged(fontFamily => {
				domDraft.findAll(`.crater-events-${property}-value`).forEach(value => {
					value.css({ fontFamily });
				});
			});

			let paneColorCell = pane.find('#color-cell').parentNode;
			this.pickColor({ parent: paneColorCell, cell: paneColorCell.find('#color-cell') }, (color) => {
				domDraft.findAll(`.crater-events-${property}-value`).forEach(value => {
					value.css({ color });
				});
				paneColorCell.find('#color-cell').value = color;
				paneColorCell.find('#color-cell').setAttribute('value', color);
			});

			pane.find('#show-cell').onChanged(display => {
				domDraft.findAll(`.crater-events-${property}`).forEach(aProperty => {
					if (display.toLowerCase() == 'no') aProperty.css({ display: 'none' });
					else aProperty.cssRemove(['display']);
				});
			});
		};

		let tasksPane = this.paneContent.find('.tasks-pane');
		handleProperties('task', tasksPane);

		let durationsPane = this.paneContent.find('.durations-pane');
		handleProperties('duration', durationsPane);

		let locationsPane = this.paneContent.find('.locations-pane');
		handleProperties('location', locationsPane);

		let datesPane = this.paneContent.find('.dates-pane');
		handleProperties('date', datesPane);

		let settingsPane = this.paneContent.find('.settings-pane');

		settingsPane.find('#Layout-cell').onChanged();

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		//on save clicked save the webpart settings and re-render
		this.paneContent.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());
			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart			
			this.sharePoint.properties.pane.content[this.key].settings.layout = settingsPane.find('#Layout-cell').value;
		});
	}

	public update(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;
		this.paneContent = this.setUpPaneContent(params);

		let paneConnection = this.sharePoint.app.find('.crater-property-connection');

		let updateWindow = this.elementModifier.createForm({
			title: 'Setup Meta Data', attributes: { id: 'meta-data-form', class: 'form' },
			contents: {
				task: { element: 'select', attributes: { id: 'meta-data-task', name: 'Task' }, options: params.options },
				icon: { element: 'select', attributes: { id: 'meta-data-icon', name: 'Icon' }, options: params.options },
				duration: { element: 'select', attributes: { id: 'meta-data-duration', name: 'Duration' }, options: params.options },
				location: { element: 'select', attributes: { id: 'meta-data-location', name: 'Location' }, options: params.options },
				date: { element: 'select', attributes: { id: 'meta-data-date', name: 'Date' }, options: params.options }
			},
			buttons: {
				submit: { element: 'button', attributes: { id: 'update-element', class: 'btn' }, text: 'Update' },
			}
		});

		let data: any = {};
		let source: any;
		updateWindow.find('#update-element').addEventListener('click', event => {
			event.preventDefault();
			data.task = updateWindow.find('#meta-data-task').value;
			data.icon = updateWindow.find('#meta-data-icon').value;
			data.duration = updateWindow.find('#meta-data-duration').value;
			data.location = updateWindow.find('#meta-data-location').value;
			data.date = updateWindow.find('#meta-data-date').value;
			source = func.extractFromJsonArray(data, params.source);

			let newContent = this.render({ source });
			draftDom.find('.crater-events-content').innerHTML = newContent.find('.crater-events-content').innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = draftDom.outerHTML;
			this.paneContent.find('.events-pane').innerHTML = this.generatePaneContent({ events: newContent.findAll('.crater-events-row') }).innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
		});


		if (!func.isnull(paneConnection)) {
			paneConnection.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
				this.element.innerHTML = draftDom.innerHTML;

				this.element.css(draftDom.css());

				this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart
			});
		}

		return updateWindow;
	}
}

class Button extends CraterWebParts {
	public params: any;
	public elementModifier = new ElementModifier();
	public paneContent: any;
	public element: any;
	private key: any;

	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.params = params;
	}

	//create the Button webpart
	public render(params) {
		if (!func.isset(params.source)) params.source = [
			{ title: 'Facebook', link: 'https://facebook.com', image: this.sharePoint.images.append, text: 'Button1' },
			{ title: 'Twitter', link: 'https://twitter.com', image: this.sharePoint.images.append, text: 'Button2' },
			{ title: 'Youtube', link: 'https://youtube.com', image: this.sharePoint.images.append, text: 'Button3' }
		];

		let button = this.createKeyedElement({ element: 'div', attributes: { class: 'crater-button crater-component', 'data-type': 'button' } });

		let content = button.makeElement({
			element: 'div', attributes: { class: 'crater-button-content' }
		});

		for (let singleButton of params.source) {
			content.makeElement({
				element: 'a', attributes: { class: 'crater-button-single', href: singleButton.link, title: singleButton.title }, children: [
					{ element: 'img', attributes: { class: 'crater-button-icon', src: singleButton.image, alt: 'Icon' } },
					{ element: 'span', attributes: { class: 'crater-button-text' }, text: singleButton.text }
				]
			});
		}

		this.key = this.key || button.dataset['key'];
		return button;
	}

	//add functionalities ofter button has been rendered
	public rendered(params) {
		this.params = params;
		this.element = params.element;
		this.key = this.element.dataset['key'];

		let buttons = this.element.findAll('.crater-button-single');

		let imageDisplay = this.sharePoint.properties.pane.content[this.key].settings.imageDisplay;
		let imageSize = this.sharePoint.properties.pane.content[this.key].settings.imageSize;
		let fontSize = this.sharePoint.properties.pane.content[this.key].settings.fontSize;
		let fontFamily = this.sharePoint.properties.pane.content[this.key].settings.fontFamily;
		let width = this.sharePoint.properties.pane.content[this.key].settings.width;
		let height = this.sharePoint.properties.pane.content[this.key].settings.height;

		buttons.forEach(button => {
			button.css({ height, width });
			button.find('.crater-button-text').css({ fontSize, fontFamily });
			button.find('.crater-button-icon').css({ width: imageSize });
			if (imageDisplay == 'No') button.find('.crater-button-icon').hide();
			else button.find('.crater-button-icon').show();
		});
	}

	//set up Button Pane content
	public setUpPaneContent(params): any {
		this.element = params.element;
		this.key = params.element.dataset['key'];

		this.paneContent = this.elementModifier.createElement({
			element: 'div',
			attributes: { class: 'crater-property-content' }
		}).monitor();

		//check if button draft is empty
		if (this.sharePoint.properties.pane.content[this.key].draft.pane.content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.pane.content;
		}
		//check if button pane has been generated before
		else if (this.sharePoint.properties.pane.content[this.key].content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].content;
		}
		//generate new button pane
		else {
			let button = this.sharePoint.properties.pane.content[this.key].draft.dom;
			let buttons = button.findAll('.crater-button-single');

			this.paneContent.makeElement({
				element: 'div', children: [
					this.elementModifier.createElement({
						element: 'button', attributes: { class: 'btn new-component', style: { display: 'inline-block', borderRadius: '5px' } }, text: 'Add New'
					})
				]
			});

			this.paneContent.append(this.generatePaneContent({ buttons }));

			let settings = this.paneContent.makeElement({
				element: 'div', attributes: { class: 'card settings-pane' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: "Button Settings"
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'Image Size',
							}),
							this.elementModifier.cell({
								element: 'select', name: 'Image Display', options: ['Yes', 'No']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Font Size'
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Font Family'
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Width'
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Height'
							})
						]
					})
				]
			});
		}

		return this.paneContent;
	}

	private generatePaneContent(params) {
		let buttonContents = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'card content-pane' }, children: [
				{
					element: 'div', attributes: { class: 'card-title' }, children: [
						this.elementModifier.createElement({
							element: 'h2', attributes: { class: 'title' }, text: "Button Contents"
						})
					]
				}
			]
		});
		for (let element of params.buttons) {
			buttonContents.makeElement({
				element: 'div', attributes: { class: 'row button-pane' }, children: [
					this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-single-button' }),
					this.elementModifier.cell({
						element: 'img', name: 'Image', attributes: {}, dataAttributes: { class: 'crater-icon', src: element.find('.crater-button-icon').src }
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Text',
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Link'
					}),
					this.elementModifier.cell({
						element: 'input', name: 'FontColor', list: func.colors
					}),
					this.elementModifier.cell({
						element: 'input', name: 'BackgroundColor', list: func.colors
					})
				]
			});
		}

		return buttonContents;
	}

	public listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		//fetch panecontent and monitor it
		this.paneContent = this.sharePoint.app.find('.crater-property-content').monitor();

		//fetch the content of Button
		let content = this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-button-content');
		let singleButtons = content.findAll('.crater-button-single');

		//fetch the icon
		let icons = content.findAll('.crater-button-icon');

		//fetch the text
		let texts = content.findAll('.crater-button-text');

		let buttonPrototype = this.elementModifier.createElement({
			element: 'a', attributes: { class: 'crater-button-single', href: 'www.google.com', title: 'Title' }, children: [
				{ element: 'img', attributes: { class: 'crater-button-icon', src: this.sharePoint.images.append, alt: 'Icon' } },
				{ element: 'span', attributes: { class: 'crater-button-text' }, text: 'Button' }
			]
		});

		let buttonPanePrototype = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'row button-pane' }, children: [
				this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-single-button' }),
				this.elementModifier.cell({
					element: 'img', name: 'Image', attributes: {}, dataAttributes: { class: 'crater-icon', src: this.sharePoint.images.append }
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Text',
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Link'
				}),
				this.elementModifier.cell({
					element: 'input', name: 'FontColor',
				}),
				this.elementModifier.cell({
					element: 'input', name: 'BackgroundColor',
				})
			]
		});

		let buttonHandler = (buttonPane, buttonDom) => {
			//set the text of the button
			buttonPane.addEventListener('mouseover', event => {
				buttonPane.find('.crater-content-options').css({ visibility: 'visible' });
			});

			buttonPane.addEventListener('mouseout', event => {
				buttonPane.find('.crater-content-options').css({ visibility: 'hidden' });
			});

			buttonPane.find('.delete-crater-single-button').addEventListener('click', event => {
				buttonDom.remove();
				buttonPane.remove();
			});

			buttonPane.find('.add-before-crater-single-button').addEventListener('click', event => {
				let newButtonPrototype = buttonPrototype.cloneNode(true);
				let newButtonPanePrototype = buttonPanePrototype.cloneNode(true);

				buttonDom.before(newButtonPrototype);
				buttonPane.before(newButtonPanePrototype);
				buttonHandler(newButtonPanePrototype, newButtonPrototype);
			});

			buttonPane.find('.add-after-crater-single-button').addEventListener('click', event => {
				let newButtonPrototype = buttonPrototype.cloneNode(true);
				let newButtonPanePrototype = buttonPanePrototype.cloneNode(true);

				buttonDom.after(newButtonPrototype);
				buttonPane.after(newButtonPanePrototype);
				buttonHandler(newButtonPanePrototype, newButtonPrototype);
			});

			buttonPane.find('#Text-cell').onChanged(value => {
				buttonDom.find('.crater-button-text').innerText = value;
			});

			buttonPane.find('#Link-cell').onChanged(href => {
				buttonDom.setAttribute('href', href);
			});

			let colorCell = buttonPane.find('#FontColor-cell').parentNode;
			this.pickColor({ parent: colorCell, cell: colorCell.find('#FontColor-cell') }, (color) => {
				buttonDom.find('.crater-button-text').css({ color }); colorCell.find('#FontColor-cell').value = color;
				colorCell.find('#FontColor-cell').setAttribute('value', color);
			});

			let backgroundColorCell = buttonPane.find('#BackgroundColor-cell').parentNode;
			this.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.find('#BackgroundColor-cell') }, (backgroundColor) => {
				buttonDom.css({ backgroundColor });
				backgroundColorCell.find('#BackgroundColor-cell').value = backgroundColor;
				backgroundColorCell.find('#BackgroundColor-cell').setAttribute('value', backgroundColor);
			});

			let imageCell = buttonPane.find('#Image-cell').parentNode;
			//upload the icon
			this.uploadImage({ parent: imageCell }, (image) => {
				imageCell.find('#Image-cell').src = image.src;
				buttonDom.find('.crater-button-icon').src = image.src;
			});
		};

		this.paneContent.find('.new-component').addEventListener('click', event => {
			let newButtonPrototype = buttonPrototype.cloneNode(true);
			let newButtonPanePrototype = buttonPanePrototype.cloneNode(true);

			content.append(newButtonPrototype);//c
			this.paneContent.find('.content-pane').append(newButtonPanePrototype);

			buttonHandler(newButtonPanePrototype, newButtonPrototype);
		});

		this.paneContent.findAll('.button-pane').forEach((singlePane, position) => {
			buttonHandler(singlePane, singleButtons[position]);
		});

		let settingsPane = this.paneContent.find('.settings-pane');

		settingsPane.find('#Font-Size-cell').onChanged();

		settingsPane.find('#Font-Family-cell').onChanged();

		settingsPane.find('#Width-cell').onChanged();

		settingsPane.find('#Height-cell').onChanged();

		//set the display of the button
		settingsPane.find('#Image-Display-cell').onChanged();

		settingsPane.find('#Image-Size-cell').onChanged();

		// on panecontent changed set the state
		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		//on save clicked save the webpart settings and re-render
		this.paneContent.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());
			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart

			this.sharePoint.properties.pane.content[this.key].settings.ImageSize = settingsPane.find('#Image-Size-cell').value;
			this.sharePoint.properties.pane.content[this.key].settings.imageDisplay = settingsPane.find('#Image-Display-cell').value;
			this.sharePoint.properties.pane.content[this.key].settings.fontSize = settingsPane.find('#Font-Size-cell').value;
			this.sharePoint.properties.pane.content[this.key].settings.fontFamily = settingsPane.find('#Font-Family-cell').value;
			this.sharePoint.properties.pane.content[this.key].settings.width = settingsPane.find('#Width-cell').value;
			this.sharePoint.properties.pane.content[this.key].settings.height = settingsPane.find('#Height-cell').value;
		});
	}

	public update(params) {
		this.element = params.element;
		this.key = this.element.dataset.key;
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;
		this.paneContent = this.setUpPaneContent(params);

		let paneConnection = this.sharePoint.app.find('.crater-property-connection');

		let updateWindow = this.elementModifier.createForm({
			title: 'Setup Meta Data', attributes: { id: 'meta-data-form', class: 'form' },
			contents: {
				image: { element: 'select', attributes: { id: 'meta-data-image', name: 'Image' }, options: params.options },
				link: { element: 'select', attributes: { id: 'meta-data-link', name: 'Link' }, options: params.options },
				title: { element: 'select', attributes: { id: 'meta-data-title', name: 'Title' }, options: params.options, note: 'Text shown when button is hovered' },
				text: { element: 'select', attributes: { id: 'meta-data-text', name: 'Text' }, options: params.options }
			},
			buttons: {
				submit: { element: 'button', attributes: { id: 'update-element', class: 'btn' }, text: 'Update' },
			}
		});

		let data: any = {};
		let source: any;
		updateWindow.find('#update-element').addEventListener('click', event => {
			event.preventDefault();
			data.image = updateWindow.find('#meta-data-image').value;
			data.link = updateWindow.find('#meta-data-link').value;
			data.title = updateWindow.find('#meta-data-title').value;
			data.text = updateWindow.find('#meta-data-text').value;
			source = func.extractFromJsonArray(data, params.source);

			let newContent = this.render({ source });
			draftDom.find('.crater-button-content').innerHTML = newContent.find('.crater-button-content').innerHTML;

			this.sharePoint.properties.pane.content[this.key].draft.html = draftDom.outerHTML;

			this.paneContent.find('.content-pane').innerHTML = this.generatePaneContent({ buttons: newContent.findAll('.crater-button-single') }).innerHTML;

			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
		});

		if (!func.isnull(paneConnection)) {
			paneConnection.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {

				this.element.innerHTML = draftDom.innerHTML;

				this.element.css(draftDom.css());

				this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart
			});
		}

		return updateWindow;
	}
}

class Icons extends CraterWebParts {
	public params: any;
	public elementModifier = new ElementModifier();
	public paneContent: any;
	public element: any;
	private key: any;

	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.params = params;
	}

	public render(params) {
		if (!func.isset(params.source)) params.source = [
			{ title: 'Facebook', link: 'https://facebook.com', image: this.sharePoint.images.append },
			{ title: 'Twitter', link: 'https://twitter.com', image: this.sharePoint.images.append },
			{ title: 'Youtube', link: 'https://youtube.com', image: this.sharePoint.images.append }
		];

		let icons = this.createKeyedElement({ element: 'div', attributes: { class: 'crater-icons crater-component', 'data-type': 'icons' } });

		let content = icons.makeElement({
			element: 'div', attributes: { class: 'crater-icons-content' }
		});

		for (let icon of params.source) {
			content.makeElement({
				element: 'a', attributes: { class: 'crater-icons-icon crater-curve', title: icon.title, href: icon.link }, children: [
					{ element: 'img', attributes: { class: 'crater-icons-icon-image', src: icon.image, alt: icon.title } }
				]
			});
		}

		this.key = this.key || icons.dataset.key;
		return icons;
	}

	public rendered(params) {
		this.params = params;
		this.element = params.element;
		this.key = this.element.dataset['key'];

		let icons = this.element.findAll('.crater-icons-icon');
	}

	public generatePaneContent(params) {
		let iconsPane = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'card counter-pane' }, children: [
				this.elementModifier.createElement({
					element: 'div', attributes: { class: 'card-title' }, children: [
						this.elementModifier.createElement({
							element: 'h2', attributes: { class: 'title' }, text: "Icons"
						})
					]
				}),
			]
		});

		for (let i = 0; i < params.icons.length; i++) {
			iconsPane.makeElement({
				element: 'div',
				attributes: {
					style: { border: '1px solid #bbbbbb', margin: '.5em 0em' }, class: 'row crater-icons-icon-pane'
				},
				children: [
					this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-icons-icon' }),
					this.elementModifier.cell({
						element: 'img', name: 'Image', attributes: {}, dataAttributes: { class: 'crater-icon', src: params.icons[i].find('.crater-icons-icon-image').src }
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Title', value: params.icons[i].getAttribute('title')
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Link', value: params.icons[i].getAttribute('href')
					}),
				]
			});
		}

		return iconsPane;
	}

	public setUpPaneContent(params): any {
		this.element = params.element;
		this.key = params.element.dataset['key'];

		this.paneContent = this.elementModifier.createElement({
			element: 'div',
			attributes: { class: 'crater-property-content' }
		}).monitor();

		if (this.sharePoint.properties.pane.content[this.key].draft.pane.content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.pane.content;
		}
		else if (this.sharePoint.properties.pane.content[this.key].content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].content;
		}
		else {
			let icons = this.sharePoint.properties.pane.content[this.key].draft.dom.findAll('.crater-icons-icon');
			this.paneContent.makeElement({
				element: 'div', children: [
					this.elementModifier.createElement({
						element: 'button', attributes: { class: 'btn new-component', style: { display: 'inline-block', borderRadius: '5px' } }, text: 'Add New'
					})
				]
			});

			this.paneContent.append(this.generatePaneContent({ icons }));

			let settingsPane = this.paneContent.makeElement({
				element: 'div', attributes: { class: 'card settings-pane' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: "Settings"
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'Width',
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Height',
							}),
							this.elementModifier.cell({
								element: 'input', name: 'BackgroundColor', list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Color', list: func.colors
							}),
							this.elementModifier.cell({
								element: 'select', name: 'Curved', options: ['Yes', 'No']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'SpaceBetween'
							})
						]
					})
				]
			});
		}

		return this.paneContent;
	}

	public listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		this.paneContent = this.sharePoint.app.find('.crater-property-content').monitor();
		let content = this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-icons-content');
		let icons = this.sharePoint.properties.pane.content[this.key].draft.dom.findAll('.crater-icons-icon');

		let iconPanePrototype = this.elementModifier.createElement({
			element: 'div',
			attributes: {
				style: { border: '1px solid #bbbbbb', margin: '.5em 0em' }, class: 'crater-icons-icon-pane row'
			},
			children: [
				this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-icons-icon' }),
				this.elementModifier.cell({
					element: 'img', name: 'Image', attributes: {}, dataAttributes: { class: 'crater-icon', src: this.sharePoint.images.append }
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Title', value: 'Title'
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Link', value: 'Link'
				}),
			]
		});

		let iconPrototype = this.elementModifier.createElement({
			element: 'a', attributes: { class: 'crater-icons-icon crater-curve', title: 'Title', href: 'Link' }, children: [
				{ element: 'img', attributes: { class: 'crater-icons-icon-image', src: this.sharePoint.images.append, alt: 'Title' } }
			]
		});

		let iconHandler = (iconPane, iconDom) => {
			iconPane.addEventListener('mouseover', event => {
				iconPane.find('.crater-content-options').css({ visibility: 'visible' });
			});

			iconPane.addEventListener('mouseout', event => {
				iconPane.find('.crater-content-options').css({ visibility: 'hidden' });
			});

			let imageCell = iconPane.find('#Image-cell').parentNode;
			this.uploadImage({ parent: imageCell }, (image) => {
				imageCell.find('#Image-cell').src = image.src;
				iconDom.find('.crater-icons-icon-image').src = image.src;
			});

			iconPane.find('#Link-cell').onChanged(value => {
				iconDom.setAttribute('href', value);
			});

			iconPane.find('#Title-cell').onChanged(value => {
				iconDom.setAttribute('title', value);
			});

			iconPane.find('.delete-crater-icons-icon').addEventListener('click', event => {
				iconDom.remove();
				iconPane.remove();
			});

			iconPane.find('.add-before-crater-icons-icon').addEventListener('click', event => {
				let newIconPrototype = iconPrototype.cloneNode(true);
				let newColumnPanePrototype = iconPanePrototype.cloneNode(true);

				iconDom.before(newIconPrototype);
				iconPane.before(newColumnPanePrototype);
				iconHandler(newColumnPanePrototype, newIconPrototype);
			});

			iconPane.find('.add-after-crater-icons-icon').addEventListener('click', event => {
				let newIconPrototype = iconPrototype.cloneNode(true);
				let newColumnPanePrototype = iconPanePrototype.cloneNode(true);

				iconDom.after(newIconPrototype);
				iconPane.after(newColumnPanePrototype);
				iconHandler(newColumnPanePrototype, newIconPrototype);
			});
		};

		this.paneContent.find('.new-component').addEventListener('click', event => {
			let newIconPrototype = iconPrototype.cloneNode(true);
			let newIconPanePrototype = iconPanePrototype.cloneNode(true);

			content.append(newIconPrototype);//c
			this.paneContent.find('.counter-pane').append(newIconPanePrototype);

			iconHandler(newIconPanePrototype, newIconPrototype);
		});

		this.paneContent.findAll('.crater-icons-icon-pane').forEach((iconPane, position) => {
			iconHandler(iconPane, icons[position]);
		});

		this.paneContent.find('#Width-cell').onChanged();
		this.paneContent.find('#Height-cell').onChanged();
		this.paneContent.find('#SpaceBetween-cell').onChanged();
		this.paneContent.find('#Curved-cell').onChanged();


		let backgroundColorCell = this.paneContent.find('#BackgroundColor-cell').parentNode;

		this.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.find('#BackgroundColor-cell') }, (backgroundColor) => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.findAll('.crater-icons-icon').forEach(icon => {
				icon.css({
					backgroundColor
				});
			});
			backgroundColorCell.find('#BackgroundColor-cell').value = backgroundColor;
			backgroundColorCell.find('#BackgroundColor-cell').setAttribute('value', backgroundColor);
		});

		let colorCell = this.paneContent.find('#Color-cell').parentNode;

		this.pickColor({ parent: colorCell, cell: colorCell.find('#Color-cell') }, (color) => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.findAll('.crater-icons-icon').forEach(icon => {
				icon.css({
					color
				});
			});
			colorCell.find('#Color-cell').value = color;
			colorCell.find('#Color-cell').setAttribute('value', color);
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());
			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart

			this.element.findAll('.crater-icons-icon').forEach(icon => {
				if (this.paneContent.find('#Curved-cell').value == 'Yes') {
					icon.classList.add('crater-curve');
				} else {
					icon.classList.remove('crater-curve');
				}

				icon.css({
					width: this.paneContent.find('#Width-cell').value,
					height: this.paneContent.find('#Height-cell').value,
					margin: this.paneContent.find('#SpaceBetween-cell').value
				});
			});
		});
	}

	public update(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;
		this.paneContent = this.setUpPaneContent(params);

		let paneConnection = this.sharePoint.app.find('.crater-property-connection');

		let updateWindow = this.elementModifier.createForm({
			title: 'Setup Meta Data', attributes: { id: 'meta-data-form', class: 'form' },
			contents: {
				image: { element: 'select', attributes: { id: 'meta-data-image', name: 'Image' }, options: params.options },
				title: { element: 'select', attributes: { id: 'meta-data-title', name: 'Title' }, options: params.options },
				link: { element: 'select', attributes: { id: 'meta-data-link', name: 'Link' }, options: params.options },
			},
			buttons: {
				submit: { element: 'button', attributes: { id: 'update-element', class: 'btn' }, text: 'Update' },
			}
		});

		let data: any = {};
		let source: any;
		updateWindow.find('#update-element').addEventListener('click', event => {
			event.preventDefault();
			data.image = updateWindow.find('#meta-data-image').value;
			data.title = updateWindow.find('#meta-data-title').value;
			data.link = updateWindow.find('#meta-data-link').value;
			source = func.extractFromJsonArray(data, params.source);

			let newContent = this.render({ source });
			draftDom.find('.crater-icons-content').innerHTML = newContent.find('.crater-icons-content').innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = draftDom.outerHTML;

			this.paneContent.find('.counter-pane').innerHTML = this.generatePaneContent({ icons: newContent.findAll('.crater-icons-icon') }).innerHTML;

			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
		});


		if (!func.isnull(paneConnection)) {
			paneConnection.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
				this.element.innerHTML = draftDom.innerHTML;

				this.element.css(draftDom.css());

				this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart
			});
		}

		return updateWindow;
	}
}

class TextArea extends CraterWebParts {
	private element: any;
	private key: any;
	private params: any;
	private paneContent: any;

	constructor(params) {
		super({ sharePoint: params.sharePoint });
	}

	public render(params) {
		let textArea = this.createKeyedElement({
			element: 'div', attributes: { class: 'crater-textarea', 'data-type': 'textarea' }, children: [
				{ element: 'p', attributes: { class: 'crater-textarea-content' }, text: 'Lorem Ipsum Catre Matrium lotuim consinium' }
			]
		});

		return textArea;
	}

	public rendered(params) {
	}

	public setUpPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset.key;

		this.paneContent = this.elementModifier.createElement({
			element: 'div',
			attributes: { class: 'crater-property-content' }
		}).monitor();
		this.key = this.element.dataset.key;

		let view = this.sharePoint.properties.pane.content[this.key].settings.view;

		if (this.sharePoint.properties.pane.content[this.key].draft.pane.content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.pane.content;
		}
		else if (this.sharePoint.properties.pane.content[this.key].content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].content;
		}
		else {
			let textAreaPane = this.paneContent.makeElement({
				element: 'textarea', attributes: { class: 'crater-textarea-pane', id: this.key }
			});

			textAreaPane.innerHTML = this.element.find('.crater-textarea-content').innerHTML;
		}

		return this.paneContent;
	}

	public listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset.key;
		this.paneContent = this.sharePoint.app.find('.crater-property-content').monitor();

		let frBoxes = this.paneContent.findAll('.fr-box');

		for (let frBox of frBoxes) {
			frBox.remove();
		}

		this.paneContent.find('textarea').innerHTML = this.element.find('.crater-textarea-content').innerHTML;
		//Set the text editor
		let fr = new FroalaEditor('textarea#' + this.key);

		let remover = setInterval(() => {
			this.paneContent.childNodes.forEach(child => {
				if (child.classList.contains('fr-box')) {
					child.childNodes.forEach(pikin => {
						if (pikin.classList.contains('second-toolbar')) {
							pikin.innerHTML = '';
							clearInterval(remover);
						}
					});
				}
			});
		}, 1);

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart
			this.element.find('.crater-textarea-content').innerHTML = '';
			let children = this.paneContent.find('.fr-element').childNodes;
			for (let i of children) {
				this.element.find('.crater-textarea-content').append(i);
			}

			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());
		});
	}
}

class Section extends CraterWebParts {
	public params: any;
	public elementModifier = new ElementModifier();
	private element: any;
	private paneContent: any;
	private key: any;

	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.params = params;
	}

	public render(params) {

	}

	public rendered(params) {
		this.element = params.element;
		this.key = this.element.dataset.key;

		//get the the section's view[tabbed, straight]
		let view = this.sharePoint.properties.pane.content[this.key].settings.view;

		//fetch the section menu
		let menu = this.element.find('.crater-section-menu');

		if (!func.isnull(menu) && func.isset(this.element.dataset.view) && this.element.dataset.view == 'Tabbed') {//if menu exists and section is tabbed
			menu.findAll('li').forEach(li => {
				let found = false;
				let owner = li.dataset.owner;
				for (let keyedElement of this.element.find('.crater-section-content').findAll('.keyed-element')) {
					if (owner == keyedElement.dataset.key) {
						found = true;
						li.innerText = func.isset(keyedElement.dataset.title) ? keyedElement.dataset.title : keyedElement.dataset.type;
						break;
					}
				}
				//remove menu that the webpart has been deleted
				if (!found) li.remove();
			});

			//onmneu clicked change to the webpart
			menu.addEventListener('click', event => {
				if (event.target.nodeName == 'LI') {
					let li = event.target;
					this.element.findAll('.keyed-element').forEach(keyedElement => {
						keyedElement.classList.add('in-active');
						if (li.dataset.owner == keyedElement.dataset.key) {
							keyedElement.classList.remove('in-active');
						}
					});
				}
			});
			if (this.element.dataset.view == 'Tabbed') {
				//click the last menu
				let menuButtons = menu.findAll('li');
				menuButtons[menuButtons.length - 1].click();
			}
		}

		this.showOptions(this.element);
	}

	private generatePaneContent(params) {
		let sectionContentsPane = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'card section-contents-pane' }, children: [
				this.elementModifier.createElement({
					element: 'div', attributes: { class: 'card-title' }, children: [
						this.elementModifier.createElement({
							element: 'h2', attributes: { class: 'title' }, text: 'Section Contents'
						})
					]
				}),
			]
		});

		//set the pane for all the webparts in the section
		for (let i = 0; i < params.source.length; i++) {
			sectionContentsPane.makeElement({
				element: 'div',
				attributes: {
					style: { border: '1px solid #bbbbbb', margin: '.5em 0em' }, class: 'row crater-section-content-row-pane'
				},
				children: [
					this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-section-content-row' }),
					{ element: 'span', attributes: { class: 'crater-section-webpart-name' }, text: params.source[i].dataset.type }
				]
			});
		}

		return sectionContentsPane;
	}

	private setUpPaneContent(params) {
		this.element = params.element;
		this.paneContent = this.elementModifier.createElement({
			element: 'div',
			attributes: { class: 'crater-property-content' }
		}).monitor();
		this.key = this.element.dataset.key;

		//fetch the sections view
		let view = this.sharePoint.properties.pane.content[this.key].settings.view;

		if (this.sharePoint.properties.pane.content[this.key].draft.pane.content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.pane.content;
		}
		else if (this.sharePoint.properties.pane.content[this.key].content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].content;
		}
		else {
			this.paneContent.makeElement({
				element: 'div', children: [
					this.elementModifier.createElement({
						element: 'button', attributes: { class: 'btn new-component', style: { display: 'inline-block', borderRadius: '5px' } }, text: 'Add New'
					})
				]
			});

			let elementContents = this.element.find('.crater-section-content').findAll('.keyed-element');

			this.paneContent.append(this.generatePaneContent({ source: elementContents }));

			let settings = this.paneContent.makeElement({
				element: 'div', attributes: { class: 'card', style: { margin: '1em', display: 'block' } }, sync: true, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Settings'
							})
						]
					}),
				]
			});
		}

		let contents = this.element.find('.crater-section-content').findAll('.keyed-element');
		this.paneContent.find('.section-contents-pane').innerHTML = this.generatePaneContent({ source: contents }).innerHTML;

		return this.paneContent;
	}

	private listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];

		this.paneContent = this.sharePoint.app.find('.crater-property-content').monitor();

		let sectionContents = this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-section-content');

		let sectionContentDom = sectionContents.childNodes;
		let sectionContentPane = this.paneContent.find('.section-contents-pane');

		//create section content pane prototype
		let sectionContentPanePrototype = this.elementModifier.createElement({
			element: 'div',
			attributes: {
				style: { border: '1px solid #bbbbbb', margin: '.5em 0em' }, class: 'crater-section-content-row-pane row'
			},
			children: [
				this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-section-content-row' }),
				this.elementModifier.createElement({
					element: 'span', attributes: { class: 'crater-section-webpart-name' }
				}),
			]
		});

		//set all the event listeners for the section webparts[add before & after, delete]
		let sectionContentRowHandler = (sectionContentRowPane, sectionContentRowDom) => {

			sectionContentRowPane.addEventListener('mouseover', event => {
				sectionContentRowPane.find('.crater-content-options').css({ visibility: 'visible' });
			});

			sectionContentRowPane.addEventListener('mouseout', event => {
				sectionContentRowPane.find('.crater-content-options').css({ visibility: 'hidden' });
			});

			sectionContentRowPane.find('.crater-section-webpart-name').textContent = sectionContentRowDom.dataset.type;

			sectionContentRowPane.find('.delete-crater-section-content-row').addEventListener('click', event => {
				sectionContentRowDom.remove();
				sectionContentRowPane.remove();
			});

			sectionContentRowPane.find('.add-before-crater-section-content-row').addEventListener('click', event => {
				this.paneContent.append(
					this.sharePoint.displayPanel(webpart => {
						let newSectionContent = this.sharePoint.appendWebpart(sectionContents, webpart.dataset.webpart);
						sectionContentRowDom.before(newSectionContent.cloneNode(true));
						newSectionContent.remove();

						let newSectionContentRow = sectionContentPanePrototype.cloneNode(true);
						sectionContentRowPane.after(newSectionContentRow);

						sectionContentRowHandler(newSectionContentRow, newSectionContent);
					})
				);
			});

			sectionContentRowPane.find('.add-after-crater-section-content-row').addEventListener('click', event => {
				this.paneContent.append(
					this.sharePoint.displayPanel(webpart => {
						let newSectionContent = this.sharePoint.appendWebpart(sectionContents, webpart.dataset.webpart);
						sectionContentRowDom.after(newSectionContent.cloneNode(true));
						newSectionContent.remove();

						let newSectionContentRow = sectionContentPanePrototype.cloneNode(true);
						sectionContentRowPane.after(newSectionContentRow);

						sectionContentRowHandler(newSectionContentRow, newSectionContent);
					})
				);
			});
		};

		//add new webpart to the section
		this.paneContent.find('.new-component').addEventListener('click', event => {
			this.paneContent.append(
				//show the display panel and add the selected webpart
				this.sharePoint.displayPanel(webpart => {
					let newSectionContent = this.sharePoint.appendWebpart(sectionContents, webpart.dataset.webpart);
					let newSectionContentRow = sectionContentPanePrototype.cloneNode(true);
					sectionContentPane.append(newSectionContentRow);

					//listen for events on new webpart
					sectionContentRowHandler(newSectionContentRow, newSectionContent);
				})
			);
		});

		this.paneContent.findAll('.crater-section-content-row-pane').forEach((sectionContent, position) => {
			//listen for events on all webparts
			sectionContentRowHandler(sectionContent, sectionContentDom[position]);
		});

		//monitor pane contents and note the changes
		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		//save the the noted changes when save button is clicked
		this.paneContent.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;

			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());

			for (let keyedElement of this.element.findAll('.keyed-element')) {
				this[keyedElement.dataset.type]({ action: 'rendered', element: keyedElement, sharePoint: this.sharePoint });
			}
		});
	}
}

class Tab extends CraterWebParts {
	public params: any;
	public elementModifier = new ElementModifier();
	private element: any;
	private paneContent: any;
	private key: any;

	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.params = params;
	}

	public render(params) {
		let tab = this.createKeyedElement({
			element: 'div', attributes: { class: 'crater-tab crater-component crater-container', 'data-type': 'tab' }, options: ['append', 'edit', 'delete', 'clone'], children: [
				this.elementModifier.menu({ content: [] }),
				{
					element: 'div', attributes: { class: 'crater-tab-content' }
				}
			]
		});

		return tab;
	}

	public rendered(params) {
		this.element = params.element;
		this.key = this.element.dataset.key;

		//fetch the section menu
		let menu = this.element.find('.crater-menu');

		let list = [];
		let showIcons = this.sharePoint.properties.pane.content[this.key].settings.showMenuIcons;

		for (let keyedElement of this.element.find('.crater-tab-content').childNodes) {
			if (keyedElement.classList.contains('keyed-element')) {
				list.push({
					name: keyedElement.dataset.title || keyedElement.dataset.type,
					owner: keyedElement.dataset.key,
					icon: (func.isset(showIcons) && showIcons.toLowerCase() == 'yes') ? keyedElement.dataset.icon : ''
				});
			}
		}

		menu.innerHTML = this.elementModifier.menu({ content: list }).innerHTML;
		menu.css({ gridTemplateColumns: `repeat(${list.length}, 1fr)` });

		menu.findAll('.crater-menu-item-icon').forEach(icon => {
			let width = this.sharePoint.properties.pane.content[this.key].settings.iconSize || '2em';
			icon.css({ width });
		});

		//onmneu clicked change to the webpart
		menu.addEventListener('click', event => {
			if (event.target.classList.contains('crater-menu-item')) {
				let item = event.target;
				for (let keyedElement of this.element.find('.crater-tab-content').childNodes) {
					if (keyedElement.classList.contains('keyed-element')) {
						keyedElement.classList.add('in-active');
						if (item.dataset.owner == keyedElement.dataset.key) {
							keyedElement.classList.remove('in-active');
						}
					}
				}
			}
		});

		let menuButtons = menu.findAll('.crater-menu-item');
		if (func.setNotNull(menuButtons[menuButtons.length - 1])) menuButtons[menuButtons.length - 1].click();

		this.showOptions(this.element);
	}

	public generatePaneContent(params) {
		let tabContentsPane = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'card tab-contents-pane' }, children: [
				{
					element: 'div', attributes: { class: 'card-title' }, children: [
						{
							element: 'h2', attributes: { class: 'title' }, text: 'Tab Contents'
						}
					]
				},
			]
		});
		//set the pane for all the webparts in the section
		for (let i = 0; i < params.source.length; i++) {
			tabContentsPane.makeElement({
				element: 'div',
				attributes: {
					style: { border: '1px solid #bbbbbb', margin: '.5em 0em' }, class: 'row crater-tab-content-row-pane'
				},
				children: [
					this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-tab-content-row' }),
					this.elementModifier.cell({
						element: 'img', name: 'Icon', attributes: {}, dataAttributes: { class: 'crater-icon', src: params.source[i].dataset.icon || '' }
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Title', value: params.source[i].dataset.title || params.source[i].dataset.type
					})
				]
			});
		}

		return tabContentsPane;
	}

	private setUpPaneContent(params) {
		this.element = params.element;
		this.paneContent = this.elementModifier.createElement({
			element: 'div',
			attributes: { class: 'crater-property-content' }
		});

		this.key = this.element.dataset.key;
		let tab = this.sharePoint.properties.pane.content[this.key].draft.dom;

		//fetch the sections view

		if (this.sharePoint.properties.pane.content[this.key].draft.pane.content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.pane.content;
		}
		else if (this.sharePoint.properties.pane.content[this.key].content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].content;
		}
		else {
			this.paneContent.makeElement({
				element: 'div', children: [
					{
						element: 'button', attributes: { class: 'btn new-component', style: { display: 'inline-block', borderRadius: '5px' } }, text: 'Add New'
					}
				]
			});

			let menus = tab.find('.crater-menu');

			let menuPane = this.paneContent.makeElement({
				element: 'div', attributes: { class: 'menu-pane card', style: { margin: '1em', display: 'block' } }, sync: true, children: [
					{
						element: 'div', attributes: { class: 'card-title' }, children: [
							{
								element: 'h2', attributes: { class: 'title' }, text: 'Menu Settings'
							}
						]
					},
					{
						element: 'div', children: [
							this.elementModifier.cell({
								element: 'input', name: 'BackgroundColor'
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Color'
							}),
							this.elementModifier.cell({
								element: 'input', name: 'FontSize'
							}),
							this.elementModifier.cell({
								element: 'input', name: 'FontStyle'
							}),
							this.elementModifier.cell({
								element: 'select', name: 'ShowIcons', options: ['Yes', 'No']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'IconSize'
							}),
						]
					}
				]
			});

			let elementContents = tab.find('.crater-tab-content').childNodes;

			this.paneContent.append(this.generatePaneContent({ source: elementContents }));

			let settings = this.paneContent.makeElement({
				element: 'div', attributes: { class: 'card', style: { margin: '1em', display: 'block' } }, sync: true, children: [
					{
						element: 'div', attributes: { class: 'card-title' }, children: [
							{
								element: 'h2', attributes: { class: 'title' }, text: 'Settings'
							}
						]
					},
					{
						element: 'div', children: [

						]
					}
				]
			});
		}

		let contents = tab.find('.crater-tab-content').childNodes;
		this.paneContent.find('.tab-contents-pane').innerHTML = this.generatePaneContent({ source: contents }).innerHTML;

		return this.paneContent;
	}

	private listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];

		this.paneContent = this.sharePoint.app.find('.crater-property-content').monitor();

		let menu = this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-menu');
		let tabContents = this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-tab-content');

		let tabContentDom = tabContents.childNodes;
		let tabContentPane = this.paneContent.find('.tab-contents-pane');

		//fetch the current view of the section
		//create section content pane prototype
		let tabContentPanePrototype = this.elementModifier.createElement({
			element: 'div',
			attributes: {
				style: { border: '1px solid #bbbbbb', margin: '.5em 0em' }, class: 'crater-tab-content-row-pane row'
			},
			children: [
				this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-tab-content-row' }),
				this.elementModifier.cell({
					element: 'img', name: 'Icon', attributes: {}, dataAttributes: { class: 'crater-icon', src: '' }
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Title'
				})
			]
		});

		//set all the event listeners for the section webparts[add before & after, delete]
		let tabContentRowHandler = (tabContentRowPane, tabContentRowDom) => {

			tabContentRowPane.addEventListener('mouseover', event => {
				tabContentRowPane.find('.crater-content-options').css({ visibility: 'visible' });
			});

			tabContentRowPane.addEventListener('mouseout', event => {
				tabContentRowPane.find('.crater-content-options').css({ visibility: 'hidden' });
			});

			let iconCell = tabContentRowPane.find('#Icon-cell').parentNode;

			this.uploadImage({ parent: iconCell }, (image) => {
				iconCell.find('#Icon-cell').src = image.src;
				tabContentRowDom.dataset.icon = image.src;
			});

			tabContentRowPane.find('#Title-cell').onChanged(value => {
				tabContentRowDom.dataset.title = value;
			});

			tabContentRowPane.find('.delete-crater-tab-content-row').addEventListener('click', event => {
				tabContentRowDom.remove();
				tabContentRowPane.remove();
			});

			tabContentRowPane.find('.add-before-crater-tab-content-row').addEventListener('click', event => {
				this.paneContent.append(
					this.sharePoint.displayPanel(webpart => {
						let newTabContent = this.sharePoint.appendWebpart(tabContents, webpart.dataset.webpart);
						tabContentRowDom.before(newTabContent.cloneNode(true));
						newTabContent.remove();

						let newSectionContentRow = tabContentPanePrototype.cloneNode(true);
						tabContentRowPane.after(newSectionContentRow);

						tabContentRowHandler(newSectionContentRow, newTabContent);
					})
				);
			});

			tabContentRowPane.find('.add-after-crater-tab-content-row').addEventListener('click', event => {
				this.paneContent.append(
					this.sharePoint.displayPanel(webpart => {
						let newSectionContent = this.sharePoint.appendWebpart(tabContents, webpart.dataset.webpart);
						tabContentRowDom.after(newSectionContent.cloneNode(true));
						newSectionContent.remove();

						let newSectionContentRow = tabContentPanePrototype.cloneNode(true);
						tabContentRowPane.after(newSectionContentRow);

						tabContentRowHandler(newSectionContentRow, newSectionContent);
					})
				);
			});
		};

		let menuPane = this.paneContent.find('.menu-pane');

		menuPane.find('#FontSize-cell').onChanged(fontSize => {
			menu.css({ fontSize });
		});

		menuPane.find('#FontStyle-cell').onChanged(fontFamily => {
			menu.css({ fontFamily });
		});

		menuPane.find('#IconSize-cell').onChanged();
		menuPane.find('#ShowIcons-cell').onChanged();

		let backgroundColorCell = menuPane.find('#BackgroundColor-cell').parentNode;
		this.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.find('#BackgroundColor-cell') }, (backgroundColor) => {
			menu.css({ backgroundColor });
			backgroundColorCell.find('#BackgroundColor-cell').value = backgroundColor;
			backgroundColorCell.find('#BackgroundColor-cell').setAttribute('value', backgroundColor);
		});

		let colorCell = menuPane.find('#Color-cell').parentNode;
		this.pickColor({ parent: colorCell, cell: colorCell.find('#Color-cell') }, (color) => {
			menu.css({ color });
			colorCell.find('#Color-cell').value = color;
			colorCell.find('#Color-cell').setAttribute('value', color);
		});

		//add new webpart to the section
		this.paneContent.find('.new-component').addEventListener('click', event => {
			this.paneContent.append(
				//show the display panel and add the selected webpart
				this.sharePoint.displayPanel(webpart => {
					let newSectionContent = this.sharePoint.appendWebpart(tabContents, webpart.dataset.webpart);
					let newSectionContentRow = tabContentPanePrototype.cloneNode(true);
					tabContentPane.append(newSectionContentRow);

					//listen for events on new webpart
					tabContentRowHandler(newSectionContentRow, newSectionContent);
				})
			);
		});

		this.paneContent.findAll('.crater-tab-content-row-pane').forEach((sectionContent, position) => {
			//listen for events on all webparts
			tabContentRowHandler(sectionContent, tabContentDom[position]);
		});

		//monitor pane contents and note the changes
		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		//save the the noted changes when save button is clicked
		this.paneContent.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;

			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());

			for (let keyedElement of this.element.findAll('.keyed-element')) {
				this[keyedElement.dataset.type]({ action: 'rendered', element: keyedElement, sharePoint: this.sharePoint });
				console.log(keyedElement);

			}

			this.sharePoint.properties.pane.content[this.key].settings.showMenuIcons = menuPane.find('#ShowIcons-cell').value;
			this.sharePoint.properties.pane.content[this.key].settings.iconSize = menuPane.find('#IconSize-cell').value;
		});
	}
}

class Slider extends CraterWebParts {
	public params: any;
	public elementModifier = new ElementModifier();
	private element: any;
	private key: any;
	private paneContent: any;

	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.params = params;
	}

	public render(params) {
		if (!func.isset(params.source)) params.source = [
			{ image: 'https://ipigroup.sharepoint.com/sites/ShortPointTrial/SiteAssets/photo-1542178036-2e5efe4d8f83.jpg', text: 'text0', link: 'https://www.google.com', linkText: 'Button 0', subTitle: 'Sub Title 0' },
			{ image: 'https://ipigroup.sharepoint.com/sites/ShortPointTrial/SiteAssets/application-3426397_1920.jpg', text: 'text1', link: 'https://www.facebook.com', linkText: 'Button 1', subTitle: 'Sub Title 1' },
			{ image: 'https://ipigroup.sharepoint.com/sites/ShortPointTrial/SiteAssets/l1.jpg', text: 'text2', link: 'https://www.twitter.com', linkText: 'Button 2', subTitle: 'Sub Title 2' }
		];

		//create the slider element
		let slider = this.createKeyedElement({ element: 'div', attributes: { class: 'crater-slider crater-component', 'data-type': 'slider' } });
		let image, slides = [];

		//create the slider controllers
		let radioToggle = slider.makeElement({ element: 'span', attributes: { class: 'crater-top-right', id: 'crater-slide-controller' } });
		let slidesContainer = slider.makeElement({ element: 'div', attributes: { class: 'crater-slides' } });

		this.sharePoint.properties.pane.content[slider.dataset['key']].settings.duration = 10000;

		//add the slides
		for (let j in params.source) {
			image = params.source[j].image;

			slidesContainer.makeElement({
				element: 'div', attributes: { class: 'crater-slide' }, children: [
					{ element: 'img', attributes: { src: image, alt: 'Not Found' } },
					{
						element: 'span', attributes: { class: 'crater-slide-details' }, children: [
							{ element: 'p', text: params.source[j].text, attributes: { class: 'crater-slide-quote' } },
							{ element: 'p', text: params.source[j].subTitle, attributes: { class: 'crater-slide-sub-title' } },
							{ element: 'a', attributes: { href: params.source[j].link, class: 'crater-slide-link btn' }, text: params.source[j].linkText }
						]
					},
				]
			});
			radioToggle.makeElement({ element: 'input', attributes: { type: 'radio', class: 'crater-slide-radio-toggle' } });
		}

		// add the slide arrows
		slider.makeElement({ element: 'span', attributes: { class: 'crater-arrow crater-left-arrow' } });
		slider.makeElement({ element: 'span', attributes: { class: 'crater-arrow crater-right-arrow' } });

		this.key = this.key || slider.dataset.key;

		return slider;
	}

	public rendered(params) {
		this.element = params.element;
		this.key = params.element.dataset.key;

		let settings = this.sharePoint.properties.pane.content[this.key].settings;
		this.startSlide();

		let imageBightness = settings.imageBrightness || '50%';
		let imageBlur = settings.imageBlur || '3px';

		let filter = `brightness(${imageBightness}) blur(${imageBlur})`;

		//show controllers and arrows
		this.element.addEventListener('mouseenter', () => {
			this.element.findAll('.crater-arrow').forEach(arrow => {
				arrow.css({ visibility: 'visible' });
			});
			this.element.find('.crater-top-right').css({ visibility: 'visible' });
		});

		//hide controllers and arrows
		this.element.addEventListener('mouseleave', () => {
			this.element.findAll('.crater-arrow').forEach(arrow => {
				arrow.css({ visibility: 'hidden' });
			});
			this.element.find('.crater-top-right').css({ visibility: 'hidden' });
		});

		//make slides and images same as that of slider
		this.element.findAll('.crater-slide').forEach(slide => {
			slide.css({ height: this.element.position().height + 'px' });
			slide.findAll('img').forEach(img => {
				img.css({ height: this.element.position().height + 'px', filter });
			});
		});

		this.element.findAll('.crater-slide-quote').forEach(quote => {
			quote.css({ fontFamily: settings.textFontStyle, fontSize: settings.textFontSize, color: settings.textColor });
		});

		this.element.findAll('.crater-slide-link').forEach(link => {
			link.css({ fontFamily: settings.linkFontStyle, fontSize: settings.linkFontStyle, color: settings.linkColor, backgroundColor: settings.linkBackgroundColor, border: settings.linkBorder });

			if (settings.linkShow == 'No') {
				link.hide();
			} else {
				link.show();
			}
		});

		this.element.findAll('.crater-slide-sub-title').forEach(subTitle => {
			subTitle.css({ fontFamily: settings.subTitleFontStyle, fontSize: settings.subTitleFontStyle, color: settings.subTitleColor });

			if (settings.subTitleShow == 'No') {
				subTitle.hide();
			} else {
				subTitle.show();
			}
		});

		let alignSelf = 'center';
		let { contentLocation } = settings;
		if (contentLocation == 'Top') alignSelf = 'flex-start';
		else if (contentLocation == 'Bottom') alignSelf = 'flex-end';

		this.element.findAll('.crater-slide-details').forEach(detail => {
			detail.css({ alignSelf });
		});
	}

	//start the slider animation
	public startSlide() {
		this.key = this.element.dataset['key'];
		let controller = this.element.find('#crater-slide-controller'),
			arrows = this.element.findAll('.crater-arrow'),
			radios,
			slides = this.element.findAll('.crater-slide'),
			radio: any,
			current = 0,
			key = current;

		if (this.element.length < 1) return;

		//reset control buttons
		controller.innerHTML = '';

		//stack the first slide ontop
		for (let position = 0; position < slides.length; position++) {
			slides[position].css({ zIndex: 0 });
			if (position == 0) slides[position].css({ zIndex: 1 });
			controller.makeElement({
				element: 'input', attributes: { class: 'crater-slide-radio-toggle', type: 'radio' }
			});
		}
		radios = controller.findAll('.crater-slide-radio-toggle');

		//fading and fadeout animation
		let runFading = () => {
			if (key < 0) key = radios.length - 1;
			if (key >= radios.length) key = 0;
			for (let element of radios) {
				if (radio != element) element.checked = false;
			}
			slides[current].css({ opacity: 0, zIndex: 0 });
			slides[key].css({ opacity: 1, zIndex: 1 });

			current = key;
		};

		//move to next slide
		let keepSliding = () => {
			clearInterval(this.sharePoint.properties.pane.content[this.key].settings.animation);
			this.sharePoint.properties.pane.content[this.key].settings.animation = setInterval(() => {
				key++;
				runFading();
			}, this.sharePoint.properties.pane.content[this.key].settings.duration);
		};

		//run animation when arrow is clicked
		for (let arrow of arrows) {
			arrow.css({ marginTop: `${(this.element.position().height - arrow.position().height) / 2}px` });
			arrow.addEventListener('click', event => {
				if (arrow.classList.contains('crater-left-arrow')) {
					key--;
				}
				else if (arrow.classList.contains('crater-right-arrow')) {
					key++;
				}

				clearInterval(this.sharePoint.properties.pane.content[this.key].settings.animation);
				runFading();
			});
		}

		//run animation when a controller is clicked
		for (let position = 0; position < radios.length; position++) {
			radios[position].addEventListener('click', () => {
				clearInterval(this.sharePoint.properties.pane.content[this.key].settings.animation);
				key = position;
				runFading();
			});
		}

		//click the first controller and set the first slide
		radios[current].click();
		keepSliding();
	}

	private setUpPaneContent(params) {
		this.element = params.element;
		this.paneContent = this.elementModifier.createElement({
			element: 'div',
			attributes: { class: 'crater-property-content' }
		});

		let key = params.element.dataset['key'];
		if (this.sharePoint.properties.pane.content[key].draft.pane.content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[key].draft.pane.content;
		}
		else if (this.sharePoint.properties.pane.content[key].content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[key].content;
		}
		else {
			this.paneContent.makeElement({
				element: 'div', children: [
					this.elementModifier.createElement({
						element: 'button', attributes: { class: 'btn new-component', style: { display: 'inline-block', borderRadius: '5px' } }, text: 'Add New'
					})
				]
			});

			let slides = this.sharePoint.properties.pane.content[key].draft.dom.findAll('.crater-slide');

			this.paneContent.append(this.generatePaneContent({ slides }));

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'card text-settings' }, children: [
					{
						element: 'div', attributes: { class: 'card-title' }, children: [
							{
								element: 'h2', attributes: { class: 'title' }, text: 'Edit Text'
							}]
					},
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'Font Style', list: func.fontStyles
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Color', list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Font Size', list: func.pixelSizes
							})
						]
					})
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'card link-settings' }, children: [
					{
						element: 'div', attributes: { class: 'card-title' }, children: [
							{
								element: 'h2', attributes: { class: 'title' }, text: 'Edit Links'
							}]
					},
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'Font Style', list: func.fontStyles
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Color', list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Font Size', list: func.pixelSizes
							}),
							this.elementModifier.cell({
								element: 'select', name: 'Show', options: ['Yes', 'No']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Background Color', list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Border', list: func.borders
							})
						]
					})
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'card sub-title-settings' }, children: [
					{
						element: 'div', attributes: { class: 'card-title' }, children: [
							{
								element: 'h2', attributes: { class: 'title' }, text: 'Edit Sub-Titles'
							}]
					},
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'Font Style', list: func.fontStyles
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Color', list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Font Size', list: func.pixelSizes
							}),
							this.elementModifier.cell({
								element: 'select', name: 'Show', options: ['Yes', 'No']
							})
						]
					})
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'settings-pane card', style: { margin: '1em', display: 'block' } }, sync: true, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Settings'
							})
						]
					}),

					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'Duration', value: this.sharePoint.properties.pane.content[key].settings.duration
							}),
							this.elementModifier.cell({
								element: 'select', name: 'Content Location', options: ['Top', 'Center', 'Bottom']
							}),
							this.elementModifier.cell({
								element: 'select', name: 'View', options: ['Same Window', 'New Window', 'Pop Up']
							}),
							this.elementModifier.cell({
								element: 'select', name: 'Image Blur', options: func.pixelSizes
							}),
							this.elementModifier.cell({
								element: 'select', name: 'Image Brightness', options: func.range(0, 100)
							}),
						]
					})
				]
			});
		}

		return this.paneContent;
	}

	public generatePaneContent(params) {
		let listPane = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'card list-pane' }, children: [
				this.elementModifier.createElement({
					element: 'div', attributes: { class: 'card-title' }, children: [
						this.elementModifier.createElement({
							element: 'h2', attributes: { class: 'title' }, text: 'Slides List'
						})
					]
				}),
			]
		});

		for (let i = 0; i < params.slides.length; i++) {
			listPane.makeElement({
				element: 'div',
				attributes: {
					style: { border: '1px solid #bbbbbb', margin: '.5em 0em' }, class: 'crater-slide-row-pane row'
				},
				children: [
					this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-slide-content-row' }),
					this.elementModifier.cell({
						element: 'img', name: 'Image', attributes: {}, dataAttributes: { class: 'crater-icon', src: params.slides[i].find('img').src }
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Quote', value: params.slides[i].find('.crater-slide-quote').innerText
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Link', value: params.slides[i].find('.crater-slide-link').href
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Link Text', value: params.slides[i].find('.crater-slide-link').innerText
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Sub Title', value: params.slides[i].find('.crater-slide-sub-title').innerText
					}),
				]
			});
		}
		return listPane;
	}

	private listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		this.paneContent = this.sharePoint.app.find('.crater-property-content').monitor();

		let slides = this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-slides');

		let slideListRows = slides.findAll('.crater-slide');

		let listRowPrototype = this.elementModifier.createElement({
			element: 'div',
			attributes: {
				style: { border: '1px solid #bbbbbb', margin: '.5em 0em' }, class: 'crater-slide-row-pane row'
			},
			children: [
				this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-slide-content-row' }),
				this.elementModifier.cell({
					element: 'img', name: 'Image', attributes: {}, dataAttributes: { class: 'crater-icon', src: this.sharePoint.images.append }
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Quote', value: 'quote'
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Link', value: 'https://google.com'
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Link Text', value: 'Link'
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Sub Title', value: 'Sub Title'
				}),
			]
		});

		let slidePrototype = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'crater-slide', style: { opacity: 0 } }, children: [
				{ element: 'img', attributes: { src: 'image', alt: 'Not Found' } },
				{
					element: 'span', attributes: { class: 'crater-slide-details' }, children: [
						{ element: 'p', text: 'Qoute', attributes: { class: 'crater-slide-quote' } },
						{ element: 'p', text: 'Sub Title', attributes: { class: 'crater-slide-sub-title' } },
						{ element: 'a', attributes: { href: 'https://google.com', class: 'crater-slide-link btn' }, text: 'Link' }
					]
				},
			]
		});

		let listRowHandler = (listRowPane, listRowDom) => {
			listRowPane.addEventListener('mouseover', event => {
				listRowPane.find('.crater-content-options').css({ visibility: 'visible' });
			});

			listRowPane.addEventListener('mouseout', event => {
				listRowPane.find('.crater-content-options').css({ visibility: 'hidden' });
			});

			let imageCell = listRowPane.find('#Image-cell').parentNode;
			this.uploadImage({ parent: imageCell }, (image) => {
				imageCell.find('#Image-cell').src = image.src;
				listRowDom.find('img').src = image.src;
			});

			listRowPane.find('#Quote-cell').onChanged(value => {
				listRowDom.find('.crater-slide-quote').innerHTML = value;
			});

			listRowPane.find('#Link-cell').onChanged(value => {
				listRowDom.find('.crater-slide-link').href = value;
			});

			listRowPane.find('#Link-Text-cell').onChanged(value => {
				listRowDom.find('.crater-slide-link').innerText = value;
			});

			listRowPane.find('#Sub-Title-cell').onChanged(value => {
				listRowDom.find('.crater-slide-sub-title').innerText = value;
			});

			listRowPane.find('.delete-crater-slide-content-row').addEventListener('click', event => {
				listRowDom.remove();
				listRowPane.remove();
			});

			listRowPane.find('.add-before-crater-slide-content-row').addEventListener('click', event => {
				let newSlide = slidePrototype.cloneNode(true);
				let newListRow = listRowPrototype.cloneNode(true);

				listRowDom.before(newSlide);
				listRowPane.before(newListRow);
				listRowHandler(newListRow, newSlide);
			});

			listRowPane.find('.add-after-crater-slide-content-row').addEventListener('click', event => {
				let newSlide = slidePrototype.cloneNode(true);
				let newListRow = listRowPrototype.cloneNode(true);

				listRowDom.after(newSlide);
				listRowPane.after(newListRow);

				listRowHandler(newListRow, newSlide);
			});
		};

		this.paneContent.find('.new-component').addEventListener('click', event => {
			let newSlide = slidePrototype.cloneNode(true);
			let newListRow = listRowPrototype.cloneNode(true);

			slides.append(newSlide);
			this.paneContent.find('.list-pane').append(newListRow);

			listRowHandler(newListRow, newSlide);
		});

		this.paneContent.findAll('.crater-slide-row-pane').forEach((listRow, position) => {
			listRowHandler(listRow, slideListRows[position]);
		});

		this.paneContent.find('#Duration-cell').onChanged();
		this.paneContent.find('#Content-Location-cell').onChanged();
		this.paneContent.find('#View-cell').onChanged();
		this.paneContent.find('#Image-Brightness-cell').onChanged();
		this.paneContent.find('#Image-Blur-cell').onChanged();

		let textSettings = this.paneContent.find('.text-settings');
		let linkSettings = this.paneContent.find('.link-settings');
		let subTitleSettings = this.paneContent.find('.sub-title-settings');

		textSettings.find('#Font-Style-cell').onChanged();
		textSettings.find('#Color-cell').onChanged();
		textSettings.find('#Font-Size-cell').onChanged();

		linkSettings.find('#Font-Style-cell').onChanged();
		linkSettings.find('#Color-cell').onChanged();
		linkSettings.find('#Font-Size-cell').onChanged();
		linkSettings.find('#Show-cell').onChanged();
		linkSettings.find('#Background-Color-cell').onChanged();
		linkSettings.find('#Border-cell').onChanged();

		subTitleSettings.find('#Font-Style-cell').onChanged();
		subTitleSettings.find('#Color-cell').onChanged();
		subTitleSettings.find('#Font-Size-cell').onChanged();
		subTitleSettings.find('#Show-cell').onChanged();

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());

			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart

			this.sharePoint.properties.pane.content[this.key].settings.duration = this.paneContent.find('#Duration-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.view = this.paneContent.find('#View-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.imageBrightness = this.paneContent.find('#Image-Brightness-cell').value + '%';

			this.sharePoint.properties.pane.content[this.key].settings.imageBlur = this.paneContent.find('#Image-Blur-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.contentLocation = this.paneContent.find('#Content-Location-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.textFontStyle = textSettings.find('#Font-Style-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.textColor = textSettings.find('#Color-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.textFontSize = textSettings.find('#Font-Size-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.linkFontStyle = linkSettings.find('#Font-Style-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.linkColor = linkSettings.find('#Color-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.linkFontSize = linkSettings.find('#Font-Size-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.linkShow = linkSettings.find('#Show-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.linkBackgroundColor = linkSettings.find('#Background-Color-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.linkBorder = linkSettings.find('#Border-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.subTitleFontStyle = subTitleSettings.find('#Font-Style-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.subTitleColor = subTitleSettings.find('#Color-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.subTitleFontSize = subTitleSettings.find('#Font-Size-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.subTitleShow = subTitleSettings.find('#Show-cell').value;
		});
	}

	public update(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;
		this.paneContent = this.setUpPaneContent(params);

		let paneConnection = this.sharePoint.app.find('.crater-property-connection');

		let updateWindow = this.elementModifier.createForm({
			title: 'Setup Meta Data', attributes: { id: 'meta-data-form', class: 'form' },
			contents: {
				image: { element: 'select', attributes: { id: 'meta-data-image', name: 'Image' }, options: params.options },
				text: { element: 'select', attributes: { id: 'meta-data-text', name: 'Text' }, options: params.options },
			},
			buttons: {
				submit: { element: 'button', attributes: { id: 'update-element', class: 'btn' }, text: 'Update' },
			}
		});

		let data: any = {};
		let source: any;
		updateWindow.find('#update-element').addEventListener('click', event => {
			event.preventDefault();
			data.image = updateWindow.find('#meta-data-image').value;
			data.text = updateWindow.find('#meta-data-text').value;

			source = func.extractFromJsonArray(data, params.source);

			let newContent = this.render({ source });
			draftDom.find('.crater-slides').innerHTML = newContent.find('.crater-slides').innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = draftDom.outerHTML;

			this.paneContent.find('.list-pane').innerHTML = this.generatePaneContent({ slides: draftDom.findAll('.crater-slide') }).innerHTML;

			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
		});


		if (!func.isnull(paneConnection)) {
			paneConnection.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
				this.element.innerHTML = draftDom.innerHTML;

				this.element.css(draftDom.css());

				this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart
			});
		}

		return updateWindow;
	}
}

class List extends CraterWebParts {
	public params: any;
	public elementModifier = new ElementModifier();
	public paneContent: any;
	public element: any;
	public key: any;

	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.params = params;
	}

	public render(params) {
		if (!func.isset(params.source))
			params.source = [
				{ title: 'Person One', job: 'CEO', image: null, link: '#' },
				{ title: 'Person Two', job: 'Manager', image: null, link: '#' },
				{ title: 'Person Three', job: 'Founder', image: null, link: '#' },
			];

		let title = 'List';
		let people = this.createKeyedElement({ element: 'div', attributes: { class: 'crater-list crater-component', 'data-type': 'list' } });

		people.makeElement({
			element: 'div', attributes: { class: 'crater-list-title' }, children: [
				{ element: 'img', attributes: { class: 'crater-list-title-icon', src: this.sharePoint.images.append } },
				{ element: 'p', text: title }
			]
		});

		let content = people.makeElement({ element: 'div', attributes: { class: 'crater-list-content' } });

		for (let person of params.source) {
			content.append(this.elementModifier.createElement({
				element: 'div', attributes: { class: 'crater-list-content-row' }, children: [
					this.elementModifier.createElement({ element: 'img', attributes: { class: 'crater-list-content-row-image', id: 'image', src: func.isnull(person.image) ? this.sharePoint.images.append : person.image, alt: "DP" } }),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'crater-list-content-row-details' }, children: [
							this.elementModifier.createElement({ element: 'a', attributes: { class: 'crater-list-content-row-details-title', id: 'title' }, text: person.title }),
							this.elementModifier.createElement({ element: 'a', attributes: { class: 'crater-list-content-row-details-job', id: 'job' }, text: person.job }),
							this.elementModifier.createElement({ element: 'a', attributes: { href: person.link, class: 'crater-list-content-row-details-link', id: 'link' }, text: 'Click for more info...' })
						]
					})
				]
			}));
		}
		return people;
	}

	public rendered(params) {

	}

	public setUpPaneContent(params): any {
		let key = params.element.dataset['key'];
		this.element = this.sharePoint.properties.pane.content[key].draft.dom;
		this.paneContent = this.elementModifier.createElement({
			element: 'div',
			attributes: { class: 'crater-property-content' }
		}).monitor();

		if (this.sharePoint.properties.pane.content[key].draft.pane.content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[key].draft.pane.content;
		}
		else if (this.sharePoint.properties.pane.content[key].content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[key].content;
		}
		else {
			let peopleList = this.sharePoint.properties.pane.content[key].draft.dom.find('.crater-list-content');
			let peopleListRows = peopleList.findAll('.crater-list-content-row');
			this.paneContent.makeElement({
				element: 'div', children: [
					this.elementModifier.createElement({
						element: 'button', attributes: { class: 'btn new-component', style: { display: 'inline-block', borderRadius: '5px' } }, text: 'Add New'
					})
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'title-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'People Title'
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'img', name: 'icon', attributes: {}, dataAttributes: { class: 'crater-icon', src: this.element.find('.crater-list-title-icon').src }
							}),
							this.elementModifier.cell({
								element: 'input', name: 'title', value: this.element.find('.crater-list-title').textContent
							}),
							this.elementModifier.cell({
								element: 'input', name: 'backgroundcolor', value: this.element.find('.crater-list-title').css()['background-color'], list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'color', value: this.element.find('.crater-list-title').css().color, list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontsize', value: this.element.find('.crater-list-title').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'height', value: this.element.find('.crater-list-title').css()['height'] || ''
							}),
							this.elementModifier.cell({
								element: 'select', name: 'Show', options: ['Yes', 'No']
							})
						]
					})
				]
			});

			this.paneContent.append(this.generatePaneContent({ list: peopleListRows }));
		}

		return this.paneContent;
	}

	public generatePaneContent(params) {
		let listPane = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'card list-pane' }, children: [
				this.elementModifier.createElement({
					element: 'div', attributes: { class: 'card-title' }, children: [
						this.elementModifier.createElement({
							element: 'h2', attributes: { class: 'title' }, text: 'People List'
						})
					]
				}),
			]
		});

		for (let i = 0; i < params.list.length; i++) {
			listPane.makeElement({
				element: 'div',
				attributes: { class: 'crater-list-row-pane row' },
				children: [
					this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-list-content-row' }),
					this.elementModifier.cell({
						element: 'img', name: 'Image', attributes: {}, dataAttributes: { class: 'crater-icon', src: params.list[i].find('#image').src }
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Title', value: params.list[i].find('#title').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Job', value: params.list[i].find('#job').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Link', value: params.list[i].find('#link').href
					}),
				]
			});
		}

		return listPane;
	}

	public listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		this.paneContent = this.sharePoint.app.find('.crater-property-content').monitor();

		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;

		let peopleList = draftDom.find('.crater-list-content');
		let peopleListRows = peopleList.findAll('.crater-list-content-row');

		let listRowPrototype = this.elementModifier.createElement({
			element: 'div',
			attributes: {
				class: 'crater-list-row-pane row'
			},
			children: [
				this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-list-content-row' }),
				this.elementModifier.cell({
					element: 'img', name: 'Image', dataAttributes: { src: this.sharePoint.images.append, class: 'crater-icon' }
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Title', value: 'title'
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Job', value: 'Job'
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Link', value: 'Link'
				}),
			]
		});

		let peopleContentRowPrototype = this.createKeyedElement({
			element: 'div', attributes: { class: 'crater-list-content-row' }, children: [
				this.elementModifier.createElement({ element: 'img', attributes: { class: 'crater-list-content-row-image', id: 'image', src: this.sharePoint.images.append, alt: 'DP' } }),
				this.elementModifier.createElement({
					element: 'div', attributes: { class: 'crater-list-content-row-details' }, children: [
						this.elementModifier.createElement({ element: 'a', attributes: { class: 'crater-list-content-row-details-title', id: 'title' }, text: 'Title' }),
						this.elementModifier.createElement({ element: 'a', attributes: { class: 'crater-list-content-row-details-job', id: 'job' }, text: 'Job' }),
						this.elementModifier.createElement({ href: '#', element: 'a', attributes: { class: 'crater-list-content-row-details-link' }, text: 'Click for more info...' })
					]
				})
			]
		});

		let listRowHandler = (listRowPane, listRowDom) => {
			listRowPane.addEventListener('mouseover', event => {
				listRowPane.find('.crater-content-options').css({ visibility: 'visible' });
			});

			listRowPane.addEventListener('mouseout', event => {
				listRowPane.find('.crater-content-options').css({ visibility: 'hidden' });
			});

			let imageCell = listRowPane.find('#Image-cell').parentNode;
			this.uploadImage({ parent: imageCell }, (image) => {
				imageCell.find('#Image-cell').src = image.src;
				listRowDom.find('.crater-list-content-row-image').src = image.src;
			});

			listRowPane.find('#Title-cell').onChanged(value => {
				listRowDom.find('.crater-list-content-row-details-title').innerHTML = value;
			});

			listRowPane.find('#Job-cell').onChanged(value => {
				listRowDom.find('.crater-list-content-row-details-job').innerHTML = value;
			});

			listRowPane.find('#Link-cell').onChanged(value => {
				listRowDom.find('.crater-list-content-row-details-link').href = value;
			});

			listRowPane.find('.delete-crater-list-content-row').addEventListener('click', event => {
				listRowDom.remove();
				listRowPane.remove();
			});

			listRowPane.find('.add-before-crater-list-content-row').addEventListener('click', event => {
				let newPeopleListRow = peopleContentRowPrototype.cloneNode(true);
				let newListRow = listRowPrototype.cloneNode(true);

				listRowDom.before(newPeopleListRow);
				listRowPane.before(newListRow);
				listRowHandler(newListRow, newPeopleListRow);
			});

			listRowPane.find('.add-after-crater-list-content-row').addEventListener('click', event => {
				let newPeopleListRow = peopleContentRowPrototype.cloneNode(true);
				let newListRow = listRowPrototype.cloneNode(true);

				listRowDom.after(newPeopleListRow);
				listRowPane.after(newListRow);

				listRowHandler(newListRow, newPeopleListRow);
			});
		};

		let titlePane = this.paneContent.find('.title-pane');
		let iconCell = titlePane.find('#icon-cell').parentNode;

		this.uploadImage({ parent: iconCell }, (image) => {
			iconCell.find('#icon-cell').src = image.src;
			this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-list-title-icon').src = image.src;
		});

		titlePane.find('#title-cell').onChanged(value => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-list-title').innerHTML = value;
		});

		titlePane.find('#Show-cell').onChanged(value => {
			if (value == 'No') {
				this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-list-title').hide();
			} else {
				this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-list-title').show();
			}
		});

		let backgroundColorCell = titlePane.find('#backgroundcolor-cell').parentNode;
		this.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.find('#backgroundcolor-cell') }, (backgroundColor) => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-list-title').css({ backgroundColor });
			backgroundColorCell.find('#backgroundcolor-cell').value = backgroundColor;
			backgroundColorCell.find('#backgroundcolor-cell').setAttribute('value', backgroundColor);
		});

		let colorCell = titlePane.find('#color-cell').parentNode;
		this.pickColor({ parent: colorCell, cell: colorCell.find('#color-cell') }, (color) => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-list-title').css({ color });
			colorCell.find('#color-cell').value = color;
			colorCell.find('#color-cell').setAttribute('value', color);
		});

		this.paneContent.find('.title-pane').find('#fontsize-cell').onChanged(value => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-list-title').css({ fontSize: value });
		});

		this.paneContent.find('.title-pane').find('#height-cell').onChanged(value => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-list-title').css({ height: value });
		});

		this.paneContent.find('.new-component').addEventListener('click', event => {
			let newPeopleListRow = peopleContentRowPrototype.cloneNode(true);
			let newListRow = listRowPrototype.cloneNode(true);

			this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-list-content').append(newPeopleListRow);//c
			this.paneContent.find('.list-pane').append(newListRow);

			listRowHandler(newListRow, newPeopleListRow);
		});

		this.paneContent.findAll('.crater-list-row-pane').forEach((listRow, position) => {
			listRowHandler(listRow, peopleListRows[position]);
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = draftDom.innerHTML;

			this.element.css(draftDom.css());

			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart
		});
	}

	public update(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;
		this.paneContent = this.setUpPaneContent(params);

		let paneConnection = this.sharePoint.app.find('.crater-property-connection');

		let updateWindow = this.elementModifier.createForm({
			title: 'Setup Meta Data', attributes: { id: 'meta-data-form', class: 'form' },
			contents: {
				image: { element: 'select', attributes: { id: 'meta-data-image', name: 'Image' }, options: params.options },
				title: { element: 'select', attributes: { id: 'meta-data-title', name: 'Title' }, options: params.options },
				job: { element: 'select', attributes: { id: 'meta-data-job', name: 'Job' }, options: params.options },
				link: { element: 'select', attributes: { id: 'meta-data-link', name: 'Link' }, options: params.options }
			},
			buttons: {
				submit: { element: 'button', attributes: { id: 'update-element', class: 'btn' }, text: 'Update' },
			}
		});

		let data: any = {};
		let source: any;
		updateWindow.find('#update-element').addEventListener('click', event => {
			event.preventDefault();
			data.image = updateWindow.find('#meta-data-image').value;
			data.title = updateWindow.find('#meta-data-title').value;
			data.job = updateWindow.find('#meta-data-job').value;
			data.link = updateWindow.find('#meta-data-link').value;
			source = func.extractFromJsonArray(data, params.source);

			let newContent = this.render({ source });
			draftDom.find('.crater-list-content').innerHTML = newContent.find('.crater-list-content').innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = draftDom.outerHTML;

			this.paneContent.find('.list-pane').innerHTML = this.generatePaneContent({ list: newContent.findAll('.crater-list-content-row') }).innerHTML;

			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
		});


		if (!func.isnull(paneConnection)) {
			paneConnection.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
				this.element.innerHTML = draftDom.innerHTML;

				this.element.css(draftDom.css());

				this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart
			});
		}

		return updateWindow;
	}
}

class Tiles extends CraterWebParts {
	public params: any;
	public elementModifier = new ElementModifier();
	public paneContent: any;
	public element: any;
	private backgroundImage = "";
	private backgroundColor = '#999999';
	private color = 'white';
	private columns: any = 3;
	private duration: any = 10;
	private height: any = 0;
	private backgroundPosition: any = 'left';
	private backgroundWidth: any = '50%';
	private backgroundHeight: any = 'inherit';

	private key: any;

	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.backgroundImage = this.sharePoint.images.append;
		this.params = params;
	}

	public render(params) {
		if (!func.isset(params.source))
			params.source = [
				{ name: 'Tile One', about: 'Hello welcome to this place', image: this.sharePoint.images.append, color: '#333333' },
				{ name: 'Tile Two', about: 'Hello welcome to this place', image: this.sharePoint.images.append, color: '#999999' },
				{ name: 'Tile Three', about: 'Hello welcome to this place', image: this.sharePoint.images.append, color: '#666666' }
			];

		let tiles = this.createKeyedElement({ element: 'div', attributes: { class: 'crater-tiles crater-component', 'data-type': 'tiles' } });

		let content = tiles.makeElement({
			element: 'div', attributes: { class: 'crater-tiles-content' }
		});

		// add all the tiles
		for (let tile of params.source) {
			content.append(this.elementModifier.createElement({
				element: 'div', attributes: { class: 'crater-tiles-content-column', style: { backgroundColor: tile.color } }, children: [
					this.elementModifier.createElement({
						element: 'img', attributes: { class: 'crater-tiles-content-column-image', src: tile.image }
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'crater-tiles-content-column-details' }, children: [
							this.elementModifier.createElement({
								element: 'div', attributes: { class: 'crater-tiles-content-column-details-name', id: 'name' }, text: tile.name
							}),
							this.elementModifier.createElement({ element: 'a', attributes: { class: 'crater-tiles-content-column-details-about', id: 'about' }, text: tile.about }),
						]
					}),
				]
			}));
		}

		//set the webparts pre-defined settings
		this.key = this.key || tiles.dataset['key'];
		this.sharePoint.properties.pane.content[this.key].settings.columns = 3;
		this.sharePoint.properties.pane.content[this.key].settings.duration = 10;
		this.sharePoint.properties.pane.content[this.key].settings.backgroundSize = 'auto 60%';
		this.sharePoint.properties.pane.content[this.key].settings.backgroundPosition = 'center';
		return tiles;
	}

	public rendered(params) {
		this.params = params;
		this.element = params.element;
		this.key = this.element.dataset['key'];
		//at most 3 in a row
		//if not upto 3 make it to match parents width
		let tiles = this.element.findAll('.crater-tiles-content-column');
		let length: number = tiles.length;
		let currentContent;

		//fetch the settings
		this.columns = this.sharePoint.properties.pane.content[this.key].settings.columns / 1;
		this.duration = this.sharePoint.properties.pane.content[this.key].settings.duration;
		this.backgroundPosition = this.sharePoint.properties.pane.content[this.key].settings.backgroundPosition;
		this.backgroundWidth = this.sharePoint.properties.pane.content[this.key].settings.backgroundWidth;
		this.backgroundHeight = this.sharePoint.properties.pane.content[this.key].settings.backgroundHeight;

		this.height = this.element.css().height;

		this.element.findAll('.crater-tiles-content').forEach(content => {
			content.remove();
		});

		//set dimension properties and get height
		for (let i = 0; i < length; i++) {
			let tile = tiles[i];

			if (i % this.columns == 0) {
				let columns = this.columns;

				if (this.columns > length - i) {
					columns = length - i;
				}

				currentContent = this.element.makeElement({ element: 'div', attributes: { class: 'crater-tiles-content', style: { 'gridTemplateColumns': `repeat(${columns}, 1fr)` } } });
			}

			currentContent.append(tile);
			//set tile background position [left, right, center]
			let tileBackground: any = {};
			tile.find('.crater-tiles-content-column-image').cssRemove(['margin-left']);
			tile.find('.crater-tiles-content-column-image').cssRemove(['margin-right']);

			let getPosition = position => {
				if (position == 'left') return 'right';
				else if (position == 'right') return 'left';
				else return position;
			};

			let direction = getPosition(func.trem(this.backgroundPosition).toLowerCase());

			tileBackground[`margin-${direction}`] = 'auto';
			tileBackground.width = this.backgroundWidth;
			tileBackground.height = this.backgroundHeight;

			tile.find('.crater-tiles-content-column-image').css(tileBackground);

			if (!func.isset(this.height) || tile.position().height > this.height) {
				this.height = tile.position().height;
			}
			//set height to the height of the longest tile
			this.height = this.height || tile.position().height;
		}

		this.element.css({ gridTemplateRows: `repeat(${Math.ceil(length / this.columns)}, '1fr)`, height: this.height });

		this.height = func.isset(this.sharePoint.properties.pane.content[this.key].settings.height)
			? this.sharePoint.properties.pane.content[this.key].settings.height
			: this.height;

		if (func.isset(this.height) && this.height.toString().indexOf('px') == -1) this.height += 'px';

		//run animation
		for (let i = 0; i < length; i++) {
			let tile = tiles[i];
			tile.css({ height: this.height });// set the gotten height

			//reset tile view to un-hovered
			tile.find('.crater-tiles-content-column-details').classList.add('crater-tiles-content-column-details-short');

			tile.find('.crater-tiles-content-column-details').classList.remove('crater-tiles-content-column-details-full');

			//animate when hovered
			tile.addEventListener('mouseenter', event => {
				tile.find('.crater-tiles-content-column-details').classList.add('crater-tiles-content-column-details-full');
				tile.find('.crater-tiles-content-column-details').classList.remove('crater-tiles-content-column-details-short');
			});

			tile.addEventListener('mouseleave', event => {
				tile.find('.crater-tiles-content-column-details').classList.add('crater-tiles-content-column-details-short');
				tile.find('.crater-tiles-content-column-details').classList.remove('crater-tiles-content-column-details-full');
			});
		}
	}

	public setUpPaneContent(params): any {
		this.element = params.element;
		this.key = params.element.dataset['key'];

		this.paneContent = this.elementModifier.createElement({
			element: 'div',
			attributes: { class: 'crater-property-content' }
		}).monitor();

		if (this.sharePoint.properties.pane.content[this.key].draft.pane.content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.pane.content;
		}
		else if (this.sharePoint.properties.pane.content[this.key].content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].content;
		}
		else {
			let tiles = this.sharePoint.properties.pane.content[this.key].draft.dom.findAll('.crater-tiles-content-column');
			this.paneContent.makeElement({
				element: 'div', children: [
					this.elementModifier.createElement({
						element: 'button', attributes: { class: 'btn new-component', style: { display: 'inline-block', borderRadius: '5px' } }, text: 'Add New'
					})
				]
			});

			this.paneContent.append(this.generatePaneContent({ tiles }));

			let settingsPane = this.paneContent.makeElement({
				element: 'div', attributes: { class: 'card settings-pane' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: "Settings"
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'Duration', value: this.sharePoint.properties.pane.content[this.key].settings.duration
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Columns', value: this.sharePoint.properties.pane.content[this.key].settings.columns || ''
							}),
							this.elementModifier.cell({
								element: 'input', name: 'BackgroundWidth', value: this.sharePoint.properties.pane.content[this.key].settings.backgroundWidth || ''
							}),
							this.elementModifier.cell({
								element: 'input', name: 'BackgroundHeight', value: this.sharePoint.properties.pane.content[this.key].settings.backgroundHeight || ''
							}),
							this.elementModifier.cell({
								element: 'select', name: 'BackgroundPosition', options: ['Left', 'Right', 'Center']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Height', value: this.sharePoint.properties.pane.content[this.key].settings.height || ''
							})
						]
					})
				]
			});
		}

		//upload the settings
		this.paneContent.find('#Duration-cell').value = this.sharePoint.properties.pane.content[this.key].settings.duration || '';

		this.paneContent.find('#Columns-cell').value = this.sharePoint.properties.pane.content[this.key].settings.columns || '';

		this.paneContent.find('#BackgroundPosition-cell').value = this.sharePoint.properties.pane.content[this.key].settings.backgroundPosition || '';

		this.paneContent.find('#BackgroundWidth-cell').value = this.sharePoint.properties.pane.content[this.key].settings.backgroundWidth || '';

		this.paneContent.find('#BackgroundHeight-cell').value = this.sharePoint.properties.pane.content[this.key].settings.backgroundHeight || '';

		this.paneContent.find('#Height-cell').value = this.sharePoint.properties.pane.content[this.key].settings.height || '';

		return this.paneContent;
	}

	public generatePaneContent(params) {
		let tilesPane = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'card tiles-pane' }, children: [
				this.elementModifier.createElement({
					element: 'div', attributes: { class: 'card-title' }, children: [
						this.elementModifier.createElement({
							element: 'h2', attributes: { class: 'title' }, text: "Tile"
						})
					]
				}),
			]
		});

		for (let i = 0; i < params.tiles.length; i++) {
			tilesPane.makeElement({
				element: 'div',
				attributes: {
					style: { border: '1px solid #bbbbbb', margin: '.5em 0em' }, class: 'crater-tiles-content-column-pane row'
				},
				children: [
					this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-tiles-content-column' }),
					this.elementModifier.cell({
						element: 'img', name: 'Image', attributes: {}, dataAttributes: { class: 'crater-icon', src: params.tiles[i].find('.crater-tiles-content-column-image').src }
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Name', value: params.tiles[i].find('#name').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'About', value: params.tiles[i].find('#about').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Color', value: func.isset(params.tiles[i].css().color) ? params.tiles[i].css().color : this.color, list: func.colors
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Background', value: func.isset(params.tiles[i].css()['background-color']) ? params.tiles[i].css()['background-color'] : this.backgroundColor, list: func.colors
					})
				]
			});
		}
		return tilesPane;
	}

	public listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];

		this.paneContent = this.sharePoint.app.find('.crater-property-content').monitor();

		let content = this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-tiles-content');
		let tiles = content.findAll('.crater-tiles-content-column');

		let columnPanePrototype = this.elementModifier.createElement({
			element: 'div',
			attributes: {
				style: { border: '1px solid #bbbbbb', margin: '.5em 0em' }, class: 'crater-tiles-content-column-pane row'
			},
			children: [
				this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-tiles-content-column' }),
				this.elementModifier.cell({
					element: 'img', name: 'Image', attributes: {}, dataAttributes: { class: 'crater-icon', src: this.sharePoint.images.append }
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Name', value: 'Name'
				}),
				this.elementModifier.cell({
					element: 'input', name: 'About', value: 'About'
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Color', value: this.color, list: func.colors
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Background', value: this.backgroundColor, list: func.colors
				}),
			]
		});

		let columnPrototype = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'crater-tiles-content-column', style: { backgroundColor: this.color } }, children: [
				this.elementModifier.createElement({
					element: 'img', attributes: { class: 'crater-tiles-content-column-image', src: this.sharePoint.images.append }
				}),
				this.elementModifier.createElement({
					element: 'div', attributes: { class: 'crater-tiles-content-column-details' }, children: [
						this.elementModifier.createElement({
							element: 'div', attributes: { class: 'crater-tiles-content-column-details-name', id: 'name' }, text: 'Name'
						}),
						this.elementModifier.createElement({ element: 'a', attributes: { class: 'crater-tiles-content-column-details-about', id: 'about' }, text: 'About' }),
					]
				}),
			]
		});

		let tilescolumnHandler = (tilesColumnPane, tilesColumnDom) => {
			tilesColumnPane.addEventListener('mouseover', event => {
				tilesColumnPane.find('.crater-content-options').css({ visibility: 'visible' });
			});

			tilesColumnPane.addEventListener('mouseout', event => {
				tilesColumnPane.find('.crater-content-options').css({ visibility: 'hidden' });
			});

			let imageCell = tilesColumnPane.find('#Image-cell').parentNode;
			this.uploadImage({ parent: imageCell }, (image) => {
				imageCell.find('#Image-cell').src = image.src;
				tilesColumnDom.find('.crater-tiles-content-column-image').src = image.src;
			});


			tilesColumnPane.find('#Name-cell').onChanged(value => {
				tilesColumnDom.find('.crater-tiles-content-column-details-name').innerHTML = value;
			});

			tilesColumnPane.find('#About-cell').onChanged(value => {
				tilesColumnDom.find('.crater-tiles-content-column-details-about').innerHTML = value;
			});

			let colorCell = tilesColumnPane.find('#Color-cell').parentNode;
			this.pickColor({ parent: colorCell, cell: colorCell.find('#Color-cell') }, (color) => {
				tilesColumnDom.css({ color });
				colorCell.find('#Color-cell').value = color;
				colorCell.find('#Color-cell').setAttribute('value', color);
			});

			let backgroundColorCell = tilesColumnPane.find('#Background-cell').parentNode;
			this.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.find('#Background-cell') }, (backgroundColor) => {
				tilesColumnDom.css({ backgroundColor });
				backgroundColorCell.find('#Background-cell').value = backgroundColor;
				backgroundColorCell.find('#Background-cell').setAttribute('value', backgroundColor);
			});

			tilesColumnPane.find('.delete-crater-tiles-content-column').addEventListener('click', event => {
				tilesColumnDom.remove();
				tilesColumnPane.remove();
			});

			tilesColumnPane.find('.add-before-crater-tiles-content-column').addEventListener('click', event => {
				let newColumnPrototype = columnPrototype.cloneNode(true);
				let newColumnPanePrototype = columnPanePrototype.cloneNode(true);

				tilesColumnDom.before(newColumnPrototype);
				tilesColumnPane.before(newColumnPanePrototype);
				tilescolumnHandler(newColumnPanePrototype, newColumnPrototype);
			});

			tilesColumnPane.find('.add-after-crater-tiles-content-column').addEventListener('click', event => {
				let newColumnPrototype = columnPrototype.cloneNode(true);
				let newColumnPanePrototype = columnPanePrototype.cloneNode(true);

				tilesColumnDom.after(newColumnPrototype);
				tilesColumnPane.after(newColumnPanePrototype);
				tilescolumnHandler(newColumnPanePrototype, newColumnPrototype);
			});
		};

		this.paneContent.find('.new-component').addEventListener('click', event => {
			let newColumnPrototype = columnPrototype.cloneNode(true);
			let newColumnPanePrototype = columnPanePrototype.cloneNode(true);

			content.append(newColumnPrototype);//c
			this.paneContent.find('.tiles-pane').append(newColumnPanePrototype);

			tilescolumnHandler(newColumnPanePrototype, newColumnPrototype);
		});

		this.paneContent.findAll('.crater-tiles-content-column-pane').forEach((tilesColumnPane, position) => {
			tilescolumnHandler(tilesColumnPane, tiles[position]);
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
			//update webpart            

			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());

			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;

			//save the new settings
			this.sharePoint.properties.pane.content[this.key].settings.duration = this.paneContent.find('#Duration-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.columns = this.paneContent.find('#Columns-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.backgroundPosition = this.paneContent.find('#BackgroundPosition-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.backgroundHeight = this.paneContent.find('#BackgroundHeight-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.backgroundWidth = this.paneContent.find('#BackgroundWidth-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.height = this.paneContent.find('#Height-cell').value;
		});
	}

	public update(params) {
		this.element = params.element;
		this.key = this.element.dataset.key;
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;
		this.paneContent = this.setUpPaneContent(params);

		let paneConnection = this.sharePoint.app.find('.crater-property-connection');

		let updateWindow = this.elementModifier.createForm({
			title: 'Setup Meta Data', attributes: { id: 'meta-data-form', class: 'form' },
			contents: {
				image: { element: 'select', attributes: { id: 'meta-data-image', name: 'Image' }, options: params.options },
				name: { element: 'select', attributes: { id: 'meta-data-name', name: 'Name' }, options: params.options },
				about: { element: 'select', attributes: { id: 'meta-data-about', name: 'About' }, options: params.options },
				color: { element: 'select', attributes: { id: 'meta-data-color', name: 'Color' }, options: params.options }
			},
			buttons: {
				submit: { element: 'button', attributes: { id: 'update-element', class: 'btn' }, text: 'Update' },
			}
		});

		let data: any = {};
		let source: any;
		updateWindow.find('#update-element').addEventListener('click', event => {
			event.preventDefault();
			data.image = updateWindow.find('#meta-data-image').value;
			data.name = updateWindow.find('#meta-data-name').value;
			data.about = updateWindow.find('#meta-data-about').value;
			data.color = updateWindow.find('#meta-data-color').value;
			source = func.extractFromJsonArray(data, params.source);

			let newContent = this.render({ source });
			draftDom.find('.crater-tiles-content').innerHTML = newContent.find('.crater-tiles-content').innerHTML;

			this.sharePoint.properties.pane.content[this.key].draft.html = draftDom.outerHTML;

			this.paneContent.find('.tiles-pane').innerHTML = this.generatePaneContent({ tiles: newContent.findAll('.crater-tiles-content-column') }).innerHTML;

			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
		});

		if (!func.isnull(paneConnection)) {
			paneConnection.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {

				this.element.innerHTML = draftDom.innerHTML;

				this.element.css(draftDom.css());

				this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart
			});
		}

		return updateWindow;
	}
}

class Counter extends CraterWebParts {
	public params: any;
	public elementModifier = new ElementModifier();
	public paneContent: any;
	public element: any;
	private backgroundImage = '';
	private backgroundColor = '#999999';
	private color = 'white';
	private columns: number = 3;
	private duration: number = 10;
	private height: any = 0;
	private backgroundPosition: any;
	private backgroundWidth: any;
	private backgroundHeight: any;
	private key: any;

	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.backgroundImage = this.sharePoint.images.append;
		this.params = params;
	}

	public render(params) {
		if (!func.isset(params.source)) params.source = [
			{ name: 'Count One', count: 50, image: this.sharePoint.images.append, unit: '%', color: '#333333' },
			{ name: 'Count Two', count: 63, image: this.sharePoint.images.append, unit: 'eggs', color: '#999999' },
			{ name: 'Count Three', count: 633, image: this.sharePoint.images.append, unit: 'people', color: '#666666' }
		];

		let counter = this.createKeyedElement({ element: 'div', attributes: { class: 'crater-counter crater-component', 'data-type': 'counter' } });

		let content = counter.makeElement({
			element: 'div', attributes: { class: 'crater-counter-content' }
		});

		for (let count of params.source) {
			content.append(this.elementModifier.createElement({
				element: 'div', attributes: { class: 'crater-counter-content-column', style: { backgroundColor: count.color } }, children: [
					{
						element: 'span', attributes: { class: 'crater-background-filter' }
					},
					{
						element: 'img', attributes: { class: 'crater-counter-content-column-image', src: count.image }
					},
					{
						element: 'div', attributes: { class: 'crater-counter-content-column-details' }, children: [
							this.elementModifier.createElement({
								element: 'span', attributes: {
									class: 'crater-counter-content-column-details-value'
								}, children: [
									this.elementModifier.createElement({ element: 'a', attributes: { 'data-count': count.count, class: 'crater-counter-content-column-details-value-count', id: 'count' }, text: count.count }),
									this.elementModifier.createElement({ element: 'a', attributes: { class: 'crater-counter-content-column-details-value-unit', id: 'unit' }, text: count.unit })
								]
							}),
							this.elementModifier.createElement({ element: 'a', attributes: { class: 'crater-counter-content-column-details-name', id: 'name' }, text: count.name }),
						]
					}
				]
			}));
		}
		this.key = this.key || counter.dataset.key;
		//upload the pre-defined settings of the webpart
		this.sharePoint.properties.pane.content[this.key].settings.columns = 3;
		this.sharePoint.properties.pane.content[this.key].settings.duration = 10;
		this.sharePoint.properties.pane.content[this.key].settings.backgroundSize = 'auto 40%';
		this.sharePoint.properties.pane.content[this.key].settings.backgroundPosition = 'left';
		return counter;
	}

	public rendered(params) {
		this.params = params;
		this.element = params.element;
		this.key = this.element.dataset['key'];
		let settings = this.sharePoint.properties.pane.content[this.key].settings;

		let counters = this.element.findAll('.crater-counter-content-column');
		let length = counters.length;
		let currentContent;

		this.columns = this.sharePoint.properties.pane.content[this.key].settings.columns / 1;
		this.duration = this.sharePoint.properties.pane.content[this.key].settings.duration / 1;
		this.backgroundPosition = this.sharePoint.properties.pane.content[this.key].settings.backgroundPosition;
		this.backgroundWidth = this.sharePoint.properties.pane.content[this.key].settings.backgroundWidth;
		this.backgroundHeight = this.sharePoint.properties.pane.content[this.key].settings.backgroundHeight;

		this.height = this.element.css().height;

		this.element.findAll('.crater-counter-content').forEach(content => {
			content.remove();
		});

		//set the dimensions and get the height
		for (let i = 0; i < length; i++) {
			let counter = counters[i];

			if (i % this.columns == 0) {
				let columns = this.columns;

				if (this.columns > length - i) {
					columns = length - i;
				}

				currentContent = this.element.makeElement({ element: 'div', attributes: { class: 'crater-counter-content', style: { 'gridTemplateColumns': `repeat(${columns}, 1fr)`, gridGap: settings.gap } } });
			}

			currentContent.append(counter);
			counter.find('.crater-counter-content-column-image').css({ height: this.backgroundHeight, width: this.backgroundWidth, filter: `blur(${settings.backgroundFilter})` });

			if (settings.showIcons == 'No') {
				counter.find('.crater-counter-content-column-image').hide();
			} else {
				counter.find('.crater-counter-content-column-image').show();
			}

			if (this.backgroundPosition != 'Right') {
				counter.find('.crater-counter-content-column-image').css({ gridColumnStart: 1, gridRowStart: 1 });
				counter.find('.crater-counter-content-column-details').css({ gridColumnStart: 2, gridRowStart: 1 });

			} else {
				counter.find('.crater-counter-content-column-image').css({ gridColumnStart: 2, gridRowStart: 1 });
				counter.find('.crater-counter-content-column-details').css({ gridColumnStart: 1, gridRowStart: 1 });
			}

			let count = counter.find('#count').dataset['count'];

			counter.find('#count').innerHTML = 0;
			counter.find('#unit').css({ visibility: 'hidden' });

			let interval = setInterval(() => {
				counter.find('#count').innerHTML = counter.find('#count').innerHTML / 1 + 1 / 1;
				if (counter.find('#count').innerHTML == count) {
					counter.find('#unit').css({ visibility: 'unset' });
					clearInterval(interval);
				}
			}, this.duration / count);

			if (!func.isset(this.height) || counter.position().height > this.height) {
				this.height = counter.position().height;
			}
		}

		this.element.css({ gridTemplateRows: `repeat(${Math.ceil(length / this.columns)}, '1fr)`, height: this.height });

		if (this.height.toString().indexOf('px') == -1) this.height += 'px';

		if (func.isset(settings.height)) {
			this.height = settings.height;
		}

		//set the height of the counter
		for (let i = 0; i < length; i++) {
			let counter = counters[i];
			counter.css({ height: this.height });
		}
	}

	public setUpPaneContent(params): any {
		this.element = params.element;
		this.key = params.element.dataset.key;

		this.paneContent = this.elementModifier.createElement({
			element: 'div',
			attributes: { class: 'crater-property-content' }
		}).monitor();

		if (this.sharePoint.properties.pane.content[this.key].draft.pane.content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.pane.content;
		}
		else if (this.sharePoint.properties.pane.content[this.key].content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].content;
		}
		else {
			let counters = this.sharePoint.properties.pane.content[this.key].draft.dom.findAll('.crater-counter-content-column');

			this.paneContent.makeElement({
				element: 'div', children: [
					this.elementModifier.createElement({
						element: 'button', attributes: { class: 'btn new-component', style: { display: 'inline-block', borderRadius: '5px' } }, text: 'Add New'
					})
				]
			});

			this.paneContent.append(this.generatePaneContent({ counters }));

			let settingsPane = this.paneContent.makeElement({
				element: 'div', attributes: { class: 'card settings-pane' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: "Settings"
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'Duration', value: this.sharePoint.properties.pane.content[this.key].settings.duration || ''
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Columns', value: this.sharePoint.properties.pane.content[this.key].settings.columns || ''
							}),
							this.elementModifier.cell({
								element: 'input', name: 'BackgroundWidth', value: this.sharePoint.properties.pane.content[this.key].settings.backgroundWidth || '', list: func.pixelSizes
							}),
							this.elementModifier.cell({
								element: 'input', name: 'BackgroundHeight', value: this.sharePoint.properties.pane.content[this.key].settings.backgroundHeight || '', list: func.pixelSizes
							}),
							this.elementModifier.cell({
								element: 'select', name: 'BackgroundPosition', options: ['Left', 'Right']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Box Height', value: this.sharePoint.properties.pane.content[this.key].settings.height || '', list: func.pixelSizes
							}),
							this.elementModifier.cell({
								element: 'select', name: 'Show Icons', options: ['Yes', 'No']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Gap', list: func.pixelSizes
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Background Filter'
							})
						]
					})
				]
			});
		}

		this.paneContent.find('#Box-Height-cell').value = this.sharePoint.properties.pane.content[this.key].settings.height || '';

		return this.paneContent;
	}

	public generatePaneContent(params) {
		let counterPane = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'card counter-pane' }, children: [
				this.elementModifier.createElement({
					element: 'div', attributes: { class: 'card-title' }, children: [
						this.elementModifier.createElement({
							element: 'h2', attributes: { class: 'title' }, text: "Counters"
						})
					]
				}),
			]
		});

		for (let i = 0; i < params.counters.length; i++) {
			counterPane.makeElement({
				element: 'div',
				attributes: {
					style: { border: '1px solid #bbbbbb', margin: '.5em 0em' }, class: 'crater-counter-content-column-pane row'
				},
				children: [
					this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-counter-content-column' }),
					this.elementModifier.cell({
						element: 'img', name: 'Image', attributes: {}, dataAttributes: { class: 'crater-icon', src: params.counters[i].find('.crater-counter-content-column-image').src }
					}),
					this.elementModifier.cell({
						element: 'img', name: 'BackgroundImage', attributes: {}, dataAttributes: { class: 'crater-icon', src: params.counters[i].css()['background-image'] || '' }
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Count', value: params.counters[i].find('#count').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Unit', value: params.counters[i].find('#unit').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Name', value: params.counters[i].find('#name').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Color', value: func.isset(params.counters[i].css().color) ? params.counters[i].css().color : this.color, list: func.colors
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Background', value: func.isset(params.counters[i].css()['background-color']) ? params.counters[i].css()['background-color'] : this.backgroundColor, list: func.colors
					})
				]
			});
		}

		return counterPane;
	}

	public listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		this.paneContent = this.sharePoint.app.find('.crater-property-content').monitor();
		let content = this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-counter-content');
		let counters = this.sharePoint.properties.pane.content[this.key].draft.dom.findAll('.crater-counter-content-column');

		let columnPanePrototype = this.elementModifier.createElement({
			element: 'div',
			attributes: {
				style: { border: '1px solid #bbbbbb', margin: '.5em 0em' }, class: 'crater-counter-content-column-pane row'
			},
			children: [
				this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-counter-content-column' }),
				this.elementModifier.cell({
					element: 'img', name: 'Image', attributes: {}, dataAttributes: { class: 'crater-icon', src: this.sharePoint.images.append }
				}),
				this.elementModifier.cell({
					element: 'img', name: 'BackgroundImage', attributes: {}, dataAttributes: { class: 'crater-icon', src: this.sharePoint.images.append }
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Count'
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Unit'
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Name'
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Color', list: func.colors
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Background', value: this.backgroundColor, list: func.colors
				})
			]
		});

		let columnPrototype = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'crater-counter-content-column keyed-element', style: { backgroundColor: this.backgroundColor, color: this.color } }, children: [
				this.elementModifier.createElement({
					element: 'img', attributes: { class: 'crater-counter-content-column-image', src: this.sharePoint.images.append }
				}),
				this.elementModifier.createElement({
					element: 'div', attributes: { class: 'crater-counter-content-column-details' }, children: [
						this.elementModifier.createElement({
							element: 'span', attributes: { class: 'crater-counter-content-column-details-value' }, children: [
								this.elementModifier.createElement({ element: 'a', attributes: { 'data-count': 100, class: 'crater-counter-content-column-details-value-count', id: 'count' }, text: 100 }),
								this.elementModifier.createElement({ element: 'a', attributes: { class: 'crater-counter-content-column-details-value-unit', id: 'unit' }, text: 'Unit' }),
							]
						}),
						this.elementModifier.createElement({ element: 'a', attributes: { class: 'crater-counter-content-column-details-name', id: 'name' }, text: 'Name' }),
					]
				})
			]
		});

		let countercolumnHandler = (counterColumnPane, counterColumnDom) => {
			counterColumnPane.addEventListener('mouseover', event => {
				counterColumnPane.find('.crater-content-options').css({ visibility: 'visible' });
			});

			counterColumnPane.addEventListener('mouseout', event => {
				counterColumnPane.find('.crater-content-options').css({ visibility: 'hidden' });
			});

			counterColumnPane.find('#Image-cell').onChanged(value => {
				counterColumnDom.css({ backgroundImage: `url('${value}')` });
			});

			let imageCell = counterColumnPane.find('#Image-cell').parentNode;
			this.uploadImage({ parent: imageCell }, (image) => {
				imageCell.find('#Image-cell').src = image.src;
				counterColumnDom.find('.crater-counter-content-column-image').src = image.src;
			});

			let backgroundImageCell = counterColumnPane.find('#BackgroundImage-cell').parentNode;
			this.uploadImage({ parent: backgroundImageCell }, (backgroundImage) => {
				backgroundImageCell.find('#BackgroundImage-cell').src = backgroundImage.src;
				counterColumnDom.setBackgroundImage(backgroundImage.src);
			});

			counterColumnPane.find('#Count-cell').onChanged(value => {
				counterColumnDom.find('.crater-counter-content-column-details-value-count').dataset['count'] = value;
			});

			counterColumnPane.find('#Unit-cell').onChanged(value => {
				counterColumnDom.find('.crater-counter-content-column-details-value-unit').innerHTML = value;
			});

			counterColumnPane.find('#Name-cell').onChanged(value => {
				counterColumnDom.find('.crater-counter-content-column-details-name').innerHTML = value;
			});

			let colorCell = counterColumnPane.find('#Color-cell').parentNode;
			this.pickColor({ parent: colorCell, cell: colorCell.find('#Color-cell') }, (color) => {
				counterColumnDom.css({ color });
				colorCell.find('#Color-cell').value = color;
				colorCell.find('#Color-cell').setAttribute('value', color);
			});

			let backgroundColorCell = counterColumnPane.find('#Background-cell').parentNode;
			this.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.find('#Background-cell') }, (backgroundColor) => {
				counterColumnDom.css({ backgroundColor });
				backgroundColorCell.find('#Background-cell').value = backgroundColor;
				backgroundColorCell.find('#Background-cell').setAttribute('value', backgroundColor);
			});

			counterColumnPane.find('.delete-crater-counter-content-column').addEventListener('click', event => {
				counterColumnDom.remove();
				counterColumnPane.remove();
			});

			counterColumnPane.find('.add-before-crater-counter-content-column').addEventListener('click', event => {
				let newColumnPrototype = columnPrototype.cloneNode(true);
				let newColumnPanePrototype = columnPanePrototype.cloneNode(true);

				counterColumnDom.before(newColumnPrototype);
				counterColumnPane.before(newColumnPanePrototype);
				countercolumnHandler(newColumnPanePrototype, newColumnPrototype);
			});

			counterColumnPane.find('.add-after-crater-counter-content-column').addEventListener('click', event => {
				let newColumnPrototype = columnPrototype.cloneNode(true);
				let newColumnPanePrototype = columnPanePrototype.cloneNode(true);

				counterColumnDom.after(newColumnPrototype);
				counterColumnPane.after(newColumnPanePrototype);
				countercolumnHandler(newColumnPanePrototype, newColumnPrototype);
			});
		};

		this.paneContent.find('.new-component').addEventListener('click', event => {
			let newColumnPrototype = columnPrototype.cloneNode(true);
			let newColumnPanePrototype = columnPanePrototype.cloneNode(true);

			content.append(newColumnPrototype);//c
			this.paneContent.find('.counter-pane').append(newColumnPanePrototype);

			countercolumnHandler(newColumnPanePrototype, newColumnPrototype);
		});

		this.paneContent.findAll('.crater-counter-content-column-pane').forEach((counterColumnPane, position) => {
			countercolumnHandler(counterColumnPane, counters[position]);
		});

		this.paneContent.find('#Duration-cell').onChanged();
		this.paneContent.find('#Columns-cell').onChanged();
		this.paneContent.find('#Box-Height-cell').onChanged();
		this.paneContent.find('#Gap-cell').onChanged();
		this.paneContent.find('#Show-Icons-cell').onChanged();
		this.paneContent.find('#Background-Filter-cell').onChanged();
		this.paneContent.find('#BackgroundWidth-cell').onChanged();
		this.paneContent.find('#BackgroundHeight-cell').onChanged();

		this.sharePoint.properties.pane.content[this.key].settings.backgroundPosition = this.paneContent.find('#BackgroundPosition-cell').value;

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());
			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart

			this.sharePoint.properties.pane.content[this.key].settings.duration = this.paneContent.find('#Duration-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.columns = this.paneContent.find('#Columns-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.height = this.paneContent.find('#Box-Height-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.gap = this.paneContent.find('#Gap-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.showIcons = this.paneContent.find('#Show-Icons-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.backgroundFilter = this.paneContent.find('#Background-Filter-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.backgroundWidth = this.paneContent.find('#BackgroundWidth-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.backgroundHeight = this.paneContent.find('#BackgroundHeight-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.backgroundPosition = this.paneContent.find('#BackgroundPosition-cell').value;

			this.element.findAll('.crater-counter-content-column').forEach((element, position) => {
				let pane = this.paneContent.findAll('.crater-counter-content-column-pane')[position];

				if (func.isset(pane.dataset.backgroundImage)) {
					element.setBackgroundImage(pane.dataset.backgroundImage);
				}
			});
		});
	}

	public update(params) {
		this.element = params.element;
		this.key = this.element.dataset.key;
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;
		this.paneContent = this.setUpPaneContent(params);

		let paneConnection = this.sharePoint.app.find('.crater-property-connection');

		let updateWindow = this.elementModifier.createForm({
			title: 'Setup Meta Data', attributes: { id: 'meta-data-form', class: 'form' },
			contents: {
				image: { element: 'select', attributes: { id: 'meta-data-image', name: 'Image' }, options: params.options },
				name: { element: 'select', attributes: { id: 'meta-data-name', name: 'Name' }, options: params.options },
				count: { element: 'select', attributes: { id: 'meta-data-count', name: 'Count' }, options: params.options },
				unit: { element: 'select', attributes: { id: 'meta-data-unit', name: 'Unit' }, options: params.options },
				color: { element: 'select', attributes: { id: 'meta-data-color', name: 'Color' }, options: params.options }
			},
			buttons: {
				submit: { element: 'button', attributes: { id: 'update-element', class: 'btn' }, text: 'Update' },
			}
		});

		let data: any = {};
		let source: any;
		updateWindow.find('#update-element').addEventListener('click', event => {
			event.preventDefault();
			data.image = updateWindow.find('#meta-data-image').value;
			data.name = updateWindow.find('#meta-data-name').value;
			data.count = updateWindow.find('#meta-data-count').value;
			data.unit = updateWindow.find('#meta-data-unit').value;
			data.color = updateWindow.find('#meta-data-color').value;
			source = func.extractFromJsonArray(data, params.source);

			let newContent = this.render({ source });
			draftDom.find('.crater-counter-content').innerHTML = newContent.find('.crater-counter-content').innerHTML;

			this.sharePoint.properties.pane.content[this.key].draft.html = draftDom.outerHTML;

			this.paneContent.find('.counter-pane').innerHTML = this.generatePaneContent({ counters: newContent.findAll('.crater-counter-content-column') }).innerHTML;

			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
		});

		if (!func.isnull(paneConnection)) {
			paneConnection.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {

				this.element.innerHTML = draftDom.innerHTML;

				this.element.css(draftDom.css());

				this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart
			});
		}

		return updateWindow;
	}
}

class News extends CraterWebParts {
	public params: any;
	public elementModifier = new ElementModifier();
	private element: any;
	private key: any;
	private paneContent: any;

	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.params = params;
	}

	public render(params) {
		if (!func.isset(params.source)) params.source = [
			{ link: 'https://ipigroup.sharepoint.com/sites/ShortPointTrial/SiteAssets/photo-1542178036-2e5efe4d8f83.jpg', details: 'details 1' },
			{ link: 'https://ipigroup.sharepoint.com/sites/ShortPointTrial/SiteAssets/photo-1542178036-2e5efe4d8f83.jpg', details: 'details 2' },
			{ link: 'https://ipigroup.sharepoint.com/sites/ShortPointTrial/SiteAssets/photo-1542178036-2e5efe4d8f83.jpg', details: 'details 3' }
		];

		let news = this.createKeyedElement({ element: 'div', attributes: { class: 'crater-ticker crater-component', 'data-type': 'news' } });
		let title = news.makeElement({
			element: 'div', attributes: { class: 'crater-ticker-content' }, children: [
				{
					element: 'div', attributes: { class: 'crater-ticker-title' }, children: [
						{
							element: 'span', attributes: { class: 'crater-ticker-title-text' }, text: 'Company Name'
						},
						{
							element: 'span', attributes: { class: 'crater-ticker-controller' }, children: [
								{ element: 'a', attributes: { class: 'crater-arrow crater-up-arrow' } },
								{ element: 'a', attributes: { class: 'crater-arrow crater-down-arrow' } }
							]
						}
					]
				},
				{
					element: 'div', attributes: { class: 'crater-ticker-news-container' }
				}
			]
		});

		let newsContainer = news.find('.crater-ticker-news-container');

		this.sharePoint.properties.pane.content[news.dataset['key']].settings.duration = 10000;
		this.sharePoint.properties.pane.content[news.dataset['key']].settings.animationType = 'Fade';

		for (let content of params.source) {
			newsContainer.makeElement({
				element: 'a', attributes: { class: 'crater-ticker-news', href: content.link, 'data-text': content.details }, text: content.details
			});
		}

		this.key = this.key || news.dataset.key;

		return news;
	}

	public rendered(params) {
		this.element = params.element;
		this.startSlide();
		this.element.find('.crater-ticker-title').css({ height: this.element.position().height + 'px' });
	}

	public startSlide() {
		this.key = this.element.dataset['key'];

		let news = this.element.findAll('.crater-ticker-news'),
			action = this.sharePoint.properties.pane.content[this.key].settings.animationType.toLowerCase();

		if (news.length < 2) return;

		let current = 0,
			key = 0,
			title = this.element.find('.crater-ticker-title-text').position();

		for (let count = 0; count < news.length; count++) {
			news[count].innerHTML = news[count].dataset.text;
			news[count].css({ animationName: 'fade-out', });
			if (count == current) {
				news[count].css({ animationName: 'fade-in' });
			}
		}

		let runAnimation = () => {
			if (key < 0) key = news.length - 1;
			if (key >= news.length) key = 0;
			if (action == 'fade') {
				news[current].css({ animationName: 'fade-out' });
				news[key].css({ animationName: 'fade-in' });
			}
			else if (action == 'swipe') {
				news[current].css({ animationName: 'swipe-out' });
				news[key].css({ animationName: 'swipe-in' });
			}
			else if (action == 'slide') {
				news[current].css({ animationName: 'slide-out' });
				news[key].css({ animationName: 'slide-in' });
			}
			current = key;
		};

		let keepAnimating = () => {
			clearInterval(this.sharePoint.properties.pane.content[this.key].settings.animation);
			this.sharePoint.properties.pane.content[this.key].settings.animation = setInterval(() => {
				key++;
				runAnimation();
			}, this.sharePoint.properties.pane.content[this.key].settings.duration);
		};

		this.element.findAll('.crater-arrow').forEach(arrow => {
			arrow.addEventListener('click', event => {
				if (event.target.getParents('.crater-ticker') == this.element) {
					if (event.target.classList.contains('crater-down-arrow')) {
						key--;
					}
					else if (event.target.classList.contains('crater-up-arrow')) {
						key++;
					}
					clearInterval(this.sharePoint.properties.pane.content[this.key].settings.animation);
					runAnimation();
				}
			});
		});

		this.element.findAll('.crater-arrow')[current].click();

		keepAnimating();
	}

	private setUpPaneContent(params) {
		this.element = params.element;
		this.paneContent = this.elementModifier.createElement({
			element: 'div',
			attributes: { class: 'crater-property-content' }
		}).monitor();

		this.key = params.element.dataset['key'];
		if (this.sharePoint.properties.pane.content[this.key].draft.pane.content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.pane.content;
		}
		else if (this.sharePoint.properties.pane.content[this.key].content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].content;
		}
		else {
			this.paneContent.makeElement({
				element: 'div', children: [
					this.elementModifier.createElement({
						element: 'button', attributes: { class: 'btn new-component', style: { display: 'inline-block', borderRadius: '5px' } }, text: 'Add New'
					})
				]
			});

			let titlePane = this.paneContent.makeElement({
				element: 'div', attributes: { class: 'card title-pane' }, children: [
					{
						element: 'div', attributes: { class: 'card-title' }, children: [
							{ element: 'h2', attributes: { class: 'title' }, text: 'Ticker Title' }
						]
					},
					{
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'Title', value: this.element.find('.crater-ticker-title-text').innerText
							}),
							this.elementModifier.cell({
								element: 'input', name: 'BackgroundColor', value: this.element.find('.crater-ticker-title').css()['background-color'], list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'TextColor', value: this.element.find('.crater-ticker-title').css()['color'], list: func.colors
							}),
						]
					}
				]
			});

			let news = this.sharePoint.properties.pane.content[this.key].draft.dom.findAll('.crater-ticker-news');

			this.paneContent.append(this.generatePaneContent({ news }));

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'card', style: { margin: '1em', display: 'block' } }, sync: true, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Settings'
							})
						]
					}),

					this.elementModifier.createElement({
						element: 'div', children: [
							this.elementModifier.cell({
								element: 'input', name: 'Duration', value: this.sharePoint.properties.pane.content[this.key].settings.duration
							}),
							this.elementModifier.cell({
								element: 'select', name: 'Animation', options: ['Fade', 'Swipe', 'Slide'], value: this.sharePoint.properties.pane.content[this.key].settings.animationType
							}),
							this.elementModifier.cell({
								element: 'select', name: 'View', options: ['Same Window', 'New Window', 'Pop Up'], value: this.sharePoint.properties.pane.content[this.key].settings.view
							}),
						]
					})
				]
			});
		}

		return this.paneContent;
	}

	private generatePaneContent(params) {

		let newsPane = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'card news-pane' }, children: [
				this.elementModifier.createElement({
					element: 'div', attributes: { class: 'card-title' }, children: [
						this.elementModifier.createElement({
							element: 'h2', attributes: { class: 'title' }, text: 'News List'
						})
					]
				}),
			]
		});

		for (let i = 0; i < params.news.length; i++) {
			newsPane.makeElement({
				element: 'div',
				attributes: {
					style: { border: '1px solid #bbbbbb', margin: '.5em 0em', position: 'relative' }, class: 'crater-ticker-news-pane',
				},
				children: [
					this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-ticker-content-row' }),
					this.elementModifier.cell({
						element: 'input', name: 'Link', value: params.news[i].getAttribute('href')
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Text', value: params.news[i].dataset.text
					}),
				]
			});
		}

		return newsPane;
	}

	private listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		this.paneContent = this.sharePoint.app.find('.crater-property-content').monitor();
		let domDraft = this.sharePoint.properties.pane.content[this.key].draft.dom;

		let tickerNewsContainer = domDraft.find('.crater-ticker-news-container');

		let news = tickerNewsContainer.findAll('.crater-ticker-news');

		let newsPanePrototye = this.elementModifier.createElement({
			element: 'div',
			attributes: {
				style: { border: '1px solid #bbbbbb', margin: '.5em 0em' }, class: 'crater-ticker-news-pane row'
			},
			children: [
				this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-ticker-content-row' }),
				this.elementModifier.cell({
					element: 'input', name: 'Link', value: 'Link'
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Text', value: 'Text'
				}),
			]
		});

		let newsPrototype = this.elementModifier.createElement({
			element: 'a', attributes: { class: 'crater-ticker-news', 'data-text': 'Text', href: '#' }
		});

		let newsHandler = (newsPane, newsDom) => {
			newsPane.addEventListener('mouseover', event => {
				newsPane.find('.crater-content-options').css({ visibility: 'visible' });
			});

			newsPane.addEventListener('mouseout', event => {
				newsPane.find('.crater-content-options').css({ visibility: 'hidden' });
			});

			newsPane.find('#Link-cell').onChanged(value => {
				newsDom.setAttribute('href', value);
			});

			newsPane.find('#Text-cell').onChanged(value => {
				newsDom.dataset.text = value;
			});


			newsPane.find('.delete-crater-ticker-content-row').addEventListener('click', event => {
				newsDom.remove();
				newsPane.remove();
			});

			newsPane.find('.add-before-crater-ticker-content-row').addEventListener('click', event => {
				let newSlide = newsPrototype.cloneNode(true);
				let newListRow = newsPanePrototye.cloneNode(true);

				newsDom.before(newSlide);
				newsPane.before(newListRow);
				newsHandler(newListRow, newSlide);
			});

			newsPane.find('.add-after-crater-ticker-content-row').addEventListener('click', event => {
				let newSlide = newsPrototype.cloneNode(true);
				let newListRow = newsPanePrototye.cloneNode(true);

				newsDom.after(newSlide);
				newsPane.after(newListRow);

				newsHandler(newListRow, newSlide);
			});
		};

		this.paneContent.find('#Animation-cell').onChanged();
		this.paneContent.find('#Duration-cell').onChanged();
		this.paneContent.find('#View-cell').onChanged();

		this.paneContent.find('.new-component').addEventListener('click', event => {
			let newSlide = newsPrototype.cloneNode(true);
			let newListRow = newsPanePrototye.cloneNode(true);

			tickerNewsContainer.append(newSlide);//c
			this.paneContent.find('.news-pane').append(newListRow);

			newsHandler(newListRow, newSlide);
		});

		let backgroundColorCell = this.paneContent.find('.title-pane').find('#BackgroundColor-cell').parentNode;
		this.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.find('#BackgroundColor-cell') }, (backgroundColor) => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-ticker-title').css({ backgroundColor });
			backgroundColorCell.find('#BackgroundColor-cell').value = backgroundColor;
			backgroundColorCell.find('#BackgroundColor-cell').setAttribute('value', backgroundColor);
		});

		let colorCell = this.paneContent.find('.title-pane').find('#TextColor-cell').parentNode;
		this.pickColor({ parent: colorCell, cell: colorCell.find('#TextColor-cell') }, (color) => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-ticker-title').css({ color });
			colorCell.find('#TextColor-cell').value = color;
			colorCell.find('#TextColor-cell').setAttribute('value', color);
		});

		this.paneContent.find('.title-pane').find('#Title-cell').onChanged(value => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-ticker-title-text').innerText = value;
		});

		this.paneContent.findAll('.crater-ticker-news-pane').forEach((newsPane, position) => {
			newsHandler(newsPane, news[position]);
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());

			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart

			this.sharePoint.properties.pane.content[this.key].settings.duration = this.paneContent.find('#Duration-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.animationType = this.paneContent.find('#Animation-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.view = this.paneContent.find('#View-cell').value;
		});
	}

	public update(params) {
		this.element = params.element;
		this.key = this.element.dataset.key;
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;
		this.paneContent = this.setUpPaneContent(params);

		let paneConnection = this.sharePoint.app.find('.crater-property-connection');

		let updateWindow = this.elementModifier.createForm({
			title: 'Setup Meta Data', attributes: { id: 'meta-data-form', class: 'form' },
			contents: {
				link: { element: 'select', attributes: { id: 'meta-data-link', name: 'Link' }, options: params.options },
				details: { element: 'select', attributes: { id: 'meta-data-details', name: 'Details' }, options: params.options }
			},
			buttons: {
				submit: { element: 'button', attributes: { id: 'update-element', class: 'btn' }, text: 'Update' },
			}
		});

		let data: any = {};
		let source: any;
		updateWindow.find('#update-element').addEventListener('click', event => {
			event.preventDefault();
			data.link = updateWindow.find('#meta-data-link').value;
			data.details = updateWindow.find('#meta-data-details').value;
			source = func.extractFromJsonArray(data, params.source);

			let newContent = this.render({ source });
			draftDom.find('.crater-ticker-news-container').innerHTML = newContent.find('.crater-ticker-news-container').innerHTML;

			this.sharePoint.properties.pane.content[this.key].draft.html = draftDom.outerHTML;

			this.paneContent.find('.news-pane').innerHTML = this.generatePaneContent({ news: newContent.findAll('.crater-ticker-news') }).innerHTML;

			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
		});

		if (!func.isnull(paneConnection)) {
			paneConnection.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {

				this.element.innerHTML = draftDom.innerHTML;

				this.element.css(draftDom.css());

				this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart
			});
		}

		return updateWindow;
	}
}

class Crater extends CraterWebParts {
	// this is the base webpart
	public params: any;
	public elementModifier = new ElementModifier();
	public paneContent: any;
	public element: any;
	public key: any;
	private widths: any;

	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.params = params;
	}

	public render() {
		this.element = this.createKeyedElement({
			element: 'div', attributes: { class: 'crater-crater crater-component', style: { display: 'block', minHeight: '100px', width: '100%' }, 'data-type': 'crater' }, options: ['Edit', 'Delete', 'Undo', 'Redo'], children: [
				{ element: 'div', attributes: { class: 'crater-sections-container' } }
			], alignOptions: 'left',
		});
		this.key = this.element.dataset['key'];
		this.sharePoint.properties.pane.content[this.element.dataset['key']].settings.columns = 1;
		this.sharePoint.properties.pane.content[this.element.dataset['key']].settings.columnsSizes = '1fr';
		this.sharePoint.properties.pane.content[this.element.dataset['key']].settings.widths = [];
		return this.element;
	}

	public rendered(params) {
		// return;
		this.element = params.element;
		this.key = this.element.dataset.key;
		this.params = params;

		this.resetSections({ resetWidth: params.resetWidth });
		this.showOptions(this.element);
		let sections = this.element.findAll('section.crater-section');
		let currentSection: any;
		let currentSibling: any;
		let otherDirection: any;
		let craterSectionsContainer = this.element.find('.crater-sections-container');

		// return;

		let dragging = false,
			position;

		let adjustWidth = (event) => {
			//make sure that section does not exceed the bounds of the parent/crater
			// if section or siblings width is less than or equals 50 stop dragging
			if (event.clientX > this.element.position().right - 50) return;
			if (event.clientX < this.element.position().left + 50) return;
			if (currentSection.dataset.direction == 'left' && event.clientX > otherDirection)
				return;
			if (currentSection.dataset.direction == 'right' && event.clientX < otherDirection)
				return;

			let currentWidth = currentSection.position().width;
			let siblingWidth = currentSibling.position().width;

			if (func.isnull(currentSection) || func.isnull(currentSibling)) return;

			if (event.clientX > this.element.position().left - 50 && event.clientX < this.element.position().right + 50) {
				let myWidth = currentWidth;

				if (event.clientX > currentSection.position().left) {
					if (currentSection.dataset.direction == 'left') {
						myWidth = currentSection.position().right - event.clientX;
					} else {
						myWidth = event.clientX - currentSection.position().left;
					}
				}
				else {
					myWidth = currentSection.position().right - event.clientX;
				}

				let mySiblingWidth = siblingWidth + (currentWidth - myWidth);

				if (myWidth > this.element.position().width - 50) return;
				if (mySiblingWidth > this.element.position().width - 50) return;

				if (myWidth <= 50 || mySiblingWidth <= 50) return;

				this.sharePoint.properties.pane.content[this.key].settings.widths[sections.indexOf(currentSection)] = myWidth + 'px';

				this.sharePoint.properties.pane.content[this.key].settings.widths[sections.indexOf(currentSibling)] = mySiblingWidth + 'px';

				this.sharePoint.properties.pane.content[this.key].settings.columnsSizes = func.stringReplace(this.sharePoint.properties.pane.content[this.key].settings.widths.toString(), ',', ' ');

				craterSectionsContainer.css({ gridTemplateColumns: this.sharePoint.properties.pane.content[this.key].settings.columnsSizes });
			}
		};

		let adjustHeight = (event) => {
			return;
			//make sure that section does not exceed the bounds of the parent/crater
			if (func.isnull(currentSection)) return;
			if (event.clientY < this.element.position().bottom - 100 && event.clientY > this.element.position().bottom + 100) return;

			let currentHeight = currentSection.position().height;
			let difference = event.clientY - currentSection.position().bottom;
			let height = (currentHeight + difference);

			if (height <= 100) return;

			currentSection.css({ height: height + 'px' });
		};

		let dragSection = (event) => {
			this.element.removeEventListener('mousemove', adjustWidth, false);
			this.element.removeEventListener('mousemove', adjustHeight, false);
			// get the sibling to work with
			if (func.isnull(currentSection)) return;
			position = { X: event.clientX, Y: event.clientY };

			currentSibling = currentSection.dataset.direction == 'left' ? currentSection.previousSibling : currentSection.nextSibling;

			otherDirection = currentSection.dataset.direction == 'left' ? currentSection.position().right : currentSection.position().left;

			// set dragging to true
			dragging = true;
			//drag section			
			if (currentSection.dataset.orientation == 'X') {
				this.element.addEventListener('mousemove', adjustWidth, false);
			} else {
				this.element.addEventListener('mousemove', adjustHeight, false);
			}
			this.element.findAll('.crater-section').forEach(section => {
				if (section != currentSection) section.css({ visibility: 'hidden' });
			});
			this.element.css({ cursor: 'ew-resize' });
			currentSection.css({ boxShadow: 'var(--accient-shadow)' });
		};

		let stopDraggingSection = () => {
			//cancel all dragging
			this.element.removeEventListener('mousemove', adjustWidth, false);
			this.element.removeEventListener('mousemove', adjustHeight, false);
			this.element.removeEventListener('mousedown', dragSection, false);
			dragging = false;
			this.element.findAll('.crater-section').forEach(section => {
				section.css({ visibility: 'visible' });
			});
			if (!func.isnull(currentSection)) currentSection.cssRemove('box-shadow');
			this.element.css({ cursor: 'auto' });
		};

		this.element.addEventListener('mousemove', event => {
			//check if dragging
			if (dragging) return;

			if (!func.isnull(event.target.getParents('.webpart-options'))) {
				stopDraggingSection();
				return;
			}

			currentSection = event.target;
			//get current element being hovered
			if (!currentSection.classList.contains('crater-section')) currentSection = currentSection.getParents('.crater-section');
			// get the current section being hovered
			if (!func.isnull(currentSection)) currentSection.dataset.position = func.getPositionInArray(sections, currentSection);
			//set the position of the section			
			if (!func.isnull(currentSection)) {
				//if the left or right of the section is hover 
				if (event.clientX - 10 <= currentSection.position().left && currentSection.dataset.position != '0') {
					currentSection.dataset.direction = 'left';
					currentSection.dataset.orientation = 'X';
				}
				else if (event.clientX + 10 >= currentSection.position().right && currentSection.dataset.position != sections.length - 1) {
					currentSection.dataset.direction = 'right';
					currentSection.dataset.orientation = 'X';
				}
				else if (event.clientY + 10 >= currentSection.position().bottom) {
					currentSection.dataset.direction = 'bottom';
					currentSection.dataset.orientation = 'Y';
				}
				else {
					currentSection.css({ cursor: 'auto' });
					currentSection.dataset.direction = '';
					currentSection.dataset.orientation = '';
				}

				if (func.isset(currentSection.dataset.direction) && currentSection.dataset.direction != '') {
					//check for dragging if bar is hovered         
					currentSection.dataset.orientation == 'X' ? currentSection.css({ cursor: 'ew-resize' }) : currentSection.css({ cursor: 'ns-resize' });

					this.element.addEventListener('mousedown', dragSection, false);
					this.element.addEventListener('mouseup', stopDraggingSection, false);
				}
			}
		});

		this.element.addEventListener('mouseleave', event => {
			// stop all dragging
			if (!func.isnull(currentSection) && !func.isset(currentSection.dataset.orientation)) this.element.addEventListener('mouseup', stopDraggingSection, false);
		});
	}

	public generatePaneContent(params) {
		let listPane = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'card sections-pane', style: { display: 'block' } }, children: [
				this.elementModifier.createElement({
					element: 'div', attributes: { class: 'card-title' }, children: [
						this.elementModifier.createElement({
							element: 'h2', attributes: { class: 'title' }, text: 'Crater Sections'
						})
					]
				}),
			]
		});

		for (let i = 0; i < params.source.length; i++) {
			if (params.source[i].classList.contains('webpart-options')) continue;
			listPane.makeElement({
				element: 'div',
				attributes: {
					style: { border: '1px solid #bbbbbb', margin: '.5em 0em' }, class: 'crater-section-row-pane row'
				},
				children: [
					this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-content-row' }),
					this.elementModifier.cell({
						element: 'select', name: 'Scroll', options: ["Yes", "No"]
					}),
				]
			});
		}

		return listPane;
	}

	public setUpPaneContent(params) {
		this.element = params.element;
		this.key = params.element.dataset['key'];

		this.paneContent = this.elementModifier.createElement({
			element: 'div',
			attributes: { class: 'crater-property-content', 'data-property-key': this.key }
		}).monitor();

		// this.sharePoint.properties.pane.content[this.key].draft.pane.content = '';
		// this.sharePoint.properties.pane.content[this.key].content = '';
		if (this.sharePoint.properties.pane.content[this.key].draft.pane.content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.pane.content;
		}
		else if (this.sharePoint.properties.pane.content[this.key].content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].content;
		}
		else {
			let container = this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-sections-container');

			let elementContents = container.findAll('.crater-section');

			this.paneContent.append(this.generatePaneContent({ source: elementContents }));

			let settings = this.paneContent.makeElement({
				element: 'div', attributes: { class: 'card settings-pane', style: { display: 'block' } }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Settings'
							})
						]
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Columns', value: this.sharePoint.properties.pane.content[this.key].settings.columns
					}),

					this.elementModifier.cell({
						element: 'input', name: 'Columns Sizes', value: this.sharePoint.properties.pane.content[this.key].settings.columnsSizes || ''
					}),

					this.elementModifier.cell({
						element: 'select', name: 'Scroll', options: ["Yes", "No"]
					}),
				]
			});
		}

		// upload the settings
		this.paneContent.find('#Columns-cell').value = this.sharePoint.properties.pane.content[this.key].settings.columns;

		this.paneContent.find('#Columns-Sizes-cell').value = this.sharePoint.properties.pane.content[this.key].settings.columnsSizes;

		let contents = this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-sections-container').findAll('.crater-section');

		this.paneContent.find('.sections-pane').innerHTML = this.generatePaneContent({ source: contents }).innerHTML;

		return this.paneContent;
	}

	public listenPaneContent(params?) {
		this.key = params.element.dataset['key'];
		this.element = params.element;
		this.paneContent = this.sharePoint.app.find('.crater-property-content');
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;

		let sectionRowPanes = this.paneContent.findAll('.crater-section-row-pane');
		let sections = draftDom.findAll('.crater-section');

		let settingsPane = this.paneContent.find('.settings-pane');

		settingsPane.find('#Columns-Sizes-cell').onChanged();

		settingsPane.find('#Columns-cell').onChanged(value => {
			settingsPane.find('#Columns-Sizes-cell').setAttribute('value', `repeat(${value}, 1fr)`);
			settingsPane.find('#Columns-Sizes-cell').value = `repeat(${value}, 1fr)`;
		});

		settingsPane.find('#Scroll-cell').onChanged(scroll => {
			sections.forEach(section => {
				if (scroll.toLowerCase() == 'yes') {
					section.css({ overflowY: 'auto' });
				} else {
					section.cssRemove(['overflow-y']);
				}
			});
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());

			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart

			this.sharePoint.properties.pane.content[this.key].settings.columnsSizes = settingsPane.find('#Columns-Sizes-cell').value;

			if (this.sharePoint.properties.pane.content[this.key].settings.columns < this.paneContent.find('#Columns-cell').value) {
				this.sharePoint.properties.pane.content[this.key].settings.columns = this.paneContent.find('#Columns-cell').value;
				this.sharePoint.properties.pane.content[this.key].settings.resetWidth = true;

				this.resetSections({ resetWidth: true });
			}
			else if (this.sharePoint.properties.pane.content[this.key].settings.columns > this.paneContent.find('#Columns-cell').value) { //check if the columns is less than current
				alert("New number of column should be more than current");
			}
		});
	}

	private resetSections(params) {
		params = func.isset(params) ? params : {};
		let craterSectionsContainer = this.element.find('.crater-sections-container');
		let sections = craterSectionsContainer.findAll('.crater-section');
		let count = sections.length;

		let number = this.sharePoint.properties.pane.content[this.key].settings.columns - count;
		let newSections = this.createSections({ number, height: '100px' }).findAll('.crater-section');
		//copy the current contents of the sections into the newly created sections
		for (let i = 0; i < newSections.length; i++) {
			craterSectionsContainer.append(newSections[i]);
		}

		// reset count
		count = craterSectionsContainer.findAll('.crater-section').length;

		craterSectionsContainer.css({ gridTemplateColumns: `repeat(${count}, 1fr` });
		craterSectionsContainer.findAll('.crater-section').forEach((section, position) => {
			//section has been rendered
			this[section.dataset.type]({ action: 'rendered', element: section, sharePoint: this.sharePoint });
			section.css({ width: '100%' });

			this.sharePoint.properties.pane.content[this.key].settings.widths[position] = section.position().width + 'px';
		});

		craterSectionsContainer.css({ gridTemplateColumns: this.sharePoint.properties.pane.content[this.key].settings.columnsSizes });
	}

	public createSections(params) {
		let parent = this.elementModifier.createElement({
			element: 'div', options: []
		});

		for (let i = 0; i < params.number; i++) {
			//create the sections as keyed elements
			let element = this.createKeyedElement({
				element: 'section', attributes: { class: 'crater-section crater-component crater-container', 'data-type': 'section', style: { minHeight: params.height } }, options: ['Append', 'Edit', 'Delete', 'Clone'], type: 'crater-section', alignOptions: 'right', children: [
					{ element: 'div', attributes: { class: 'crater-section-content' } }
				]
			});

			parent.append(element);
		}

		return parent;
	}
}

class Table extends CraterWebParts {

	public params: any;
	public elementModifier = new ElementModifier();
	public paneContent: any;
	public element: any;
	public key: any;

	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.params = params;
	}

	public render(params) {
		let sample = { name: 'Person Two', job: 'Manager', age: '30', salary: '2000000' };
		let tableContainer = this.createKeyedElement({ element: 'div', attributes: { class: 'crater-table-container crater-component', 'data-type': 'table' } });

		if (!func.isset(params.source)) {
			params.source = [];
			for (let i = 0; i < 5; i++) {
				params.source.push(sample);
			}
		}

		let table = this.elementModifier.createTable({
			contents: params.source, rowClass: 'crater-table-row'
		});

		table.classList.add('crater-table');

		this.sharePoint.properties.pane.content[tableContainer.dataset.key].settings.sorting = this.sharePoint.properties.pane.content[tableContainer.dataset.key].settings.sorting || {};
		this.sharePoint.properties.pane.content[tableContainer.dataset.key].settings.headers = [];

		let headers = table.findAll('th');
		for (let i = 0; i < headers.length; i++) {
			this.sharePoint.properties.pane.content[tableContainer.dataset.key].settings.headers.push(headers[i].textContent);
		}
		this.key = this.key || tableContainer.dataset.key;

		tableContainer.append(table);
		return tableContainer;
	}

	public rendered(params) {

	}

	public setUpPaneContent(params): any {
		this.element = params.element;
		this.key = params.element.dataset['key'];

		this.paneContent = this.elementModifier.createElement({
			element: 'div',
			attributes: { class: 'crater-property-content' }
		}).monitor();

		if (this.sharePoint.properties.pane.content[this.key].draft.pane.content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.pane.content;
		}
		else if (this.sharePoint.properties.pane.content[this.key].content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].content;
		}
		else {
			let table = this.sharePoint.properties.pane.content[this.key].draft.dom.find('table');

			this.paneContent.append(this.generatePaneContent({ table }));

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'card table-header-settings', style: { display: 'block' } }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Table Head'
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'fontsize'
							}),
							this.elementModifier.cell({
								element: 'input', name: 'color', list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'backgroundcolor', list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'height'
							}),
							this.elementModifier.cell({
								element: 'select', name: 'show', options: ['Yes', 'No']
							})
						]
					})
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'card table-data-settings', style: { display: 'block' } }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Table Data'
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'fontsize', attributes: { type: 'number', min: 1 }, value: ''
							}),
							this.elementModifier.cell({
								element: 'input', name: 'color', attributes: { type: 'number', min: 1 }, value: '', list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'backgroundcolor', attributes: { type: 'number', min: 1 }, value: '', list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'height', attributes: { type: 'number', min: 1 }, value: ''
							})
						]
					})
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'card table-settings', style: { display: 'block' } }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Table Settings'
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'bordersize', attributes: { type: 'text' }, value: ''
							}),
							this.elementModifier.cell({
								element: 'input', name: 'bordercolor', attributes: { type: 'text' }, value: '', list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'borderstyle', attributes: { type: 'text' }, value: ''
							})
						]
					})
				]
			});
		}

		this.paneContent.childNodes.forEach(child => {
			if (child.classList.contains('crater-content-options')) {
				child.remove();
			}
		});

		this.paneContent.find('tbody').findAll('tr').forEach(tr => {
			tr.makeElement({
				element: this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-table-content-row' })
			});
		});

		return this.paneContent;
	}

	private generatePaneContent(params) {
		let tablePane = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'table-pane card' }, children: [
				params.table.cloneNode(true)
			]
		});

		if (func.isset(params.header)) {
			tablePane.find('thead').innerHTML = params.header;
		}

		tablePane.find('thead').findAll('th').forEach(th => {
			th.css({ position: 'relative' });
			let order = this.sharePoint.properties.pane.content[this.key].settings.sorting[th.dataset.name] == 1 ? 'crater-up-arrow' : 'crater-down-arrow';

			let data = this.elementModifier.createElement({
				element: 'span', attributes: { style: { width: '100%', display: 'grid', gridTemplateColumns: '80% 20%', gridGap: '1em', } }, children: [
					{
						element: 'input', attributes: { value: th.textContent }
					},
					{
						element: 'div', attributes: { class: `crater-table-sorter crater-arrow ${order}`, style: { width: '10px', height: '10px', visibility: 'hidden' } }
					}
				]
			});

			th.innerHTML = '';
			th.makeElement({
				element: this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-table-content-column' })
			});
			th.append(data);
		});

		tablePane.find('tbody').findAll('tr').forEach(tr => {
			tr.findAll('td').forEach(td => {
				let data = this.elementModifier.createElement({
					element: 'input', attributes: { value: td.textContent }
				});
				td.innerHTML = '';
				td.append(data);
			});
		});

		return tablePane;
	}

	public listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		this.paneContent = this.sharePoint.app.find('.crater-property-content').monitor();
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;
		let table = draftDom.find('table');
		let tableBody = table.find('tbody');

		let tableRows = tableBody.findAll('tr');

		let tableRowHandler = (tableRowPane, tableRowDom) => {
			tableRowPane.addEventListener('mouseover', event => {
				tableRowPane.find('.crater-content-options').css({ visibility: 'visible' });
			});

			tableRowPane.addEventListener('mouseout', event => {
				tableRowPane.find('.crater-content-options').css({ visibility: 'hidden' });
			});

			tableRowPane.find('.delete-crater-table-content-row').addEventListener('click', event => {
				tableRowDom.remove();
				tableRowPane.remove();
			});

			tableRowPane.find('.add-before-crater-table-content-row').addEventListener('click', event => {
				let newRow = tableRowDom.cloneNode(true);
				let newRowPane = tableRowPane.cloneNode(true);

				tableRowDom.before(newRow);
				tableRowPane.before(newRowPane);
				tableRowHandler(newRowPane, newRow);
			});

			tableRowPane.find('.add-after-crater-table-content-row').addEventListener('click', event => {
				let newRow = tableRowDom.cloneNode(true);
				let newRowPane = tableRowPane.cloneNode(true);

				tableRowDom.after(newRow);
				tableRowPane.after(newRowPane);
				tableRowHandler(newRowPane, newRow);
			});

			tableRowPane.findAll('td').forEach((td, position) => {
				td.find('input').onChanged(value => {
					tableRowDom.findAll('td')[position].textContent = value;
				});
			});
		};

		let dataName = 'crater-table-data-sample';//sample name

		this.paneContent.find('tbody').findAll('tr').forEach((tableRow, position) => {
			tableRowHandler(tableRow, tableRows[position]);
		});

		let tableHeadHandler = (thPane, thDom) => {
			thPane.addEventListener('mouseover', event => {
				thPane.find('.crater-content-options').css({ visibility: 'visible' });
				thPane.find('.crater-table-sorter').css({ visibility: 'visible' });
			});

			thPane.addEventListener('mouseout', event => {
				thPane.find('.crater-content-options').css({ visibility: 'hidden' });
				thPane.find('.crater-table-sorter').css({ visibility: 'hidden' });
			});

			thPane.find('.crater-table-sorter').addEventListener('click', event => {
				let order = thPane.find('.crater-table-sorter').classList.contains('crater-up-arrow') ? -1 : 1;
				let name = thPane.dataset.name.split('crater-table-data-')[1];
				let data = this.elementModifier.sortTable(table, name, order);
				let newContent = this.render({ source: data });

				this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-table').innerHTML = newContent.find('.crater-table').innerHTML;

				thPane.find('.crater-table-sorter').classList.toggle('crater-up-arrow');
				thPane.find('.crater-table-sorter').classList.toggle('crater-down-arrow');

				this.sharePoint.properties.pane.content[this.key].settings.sorting[thPane.dataset.name] = order;

				this.paneContent.find('.table-pane').innerHTML = this.generatePaneContent({ table: newContent.find('.table') }).innerHTML;

				this.paneContent.find('thead').findAll('th').forEach((_thPane, position) => {
					tableHeadHandler(_thPane, table.find('thead').findAll('th')[position]);
				});
			});

			thPane.find('input').onChanged(value => {
				let name = 'crater-table-data-' + value.toLowerCase();
				let ths = this.paneContent.find('thead').findAll('th');

				for (let sibling of ths) {
					if (sibling != thPane && sibling.dataset.name == name) {
						alert('Column already exists, Try another name');
						return;
					}
				}

				let tds = this.paneContent.findAll('td');

				for (let i in tds) {
					let td = tds[i];
					if (td.nodeName == 'TD' && td.dataset.name == thPane.dataset.name) {
						td.dataset.name = name;
						table.findAll('td')[i].dataset.name = name;
					}
				}

				thDom.textContent = value;
				thDom.dataset.name = name;
				thPane.dataset.name = name;
			});
		};

		this.paneContent.find('thead').findAll('th').forEach((thPane, position) => {
			tableHeadHandler(thPane, table.find('thead').findAll('th')[position]);
		});

		let tableBodyDataHandler = (td) => {
			td.addEventListener('mouseover', event => {
				for (let th of this.paneContent.find('thead').findAll('th')) {
					if (th.dataset.name == td.dataset.name) {
						th.find('.crater-content-options').css({ visibility: 'visible' });
					}
				}
			});

			td.addEventListener('mouseout', event => {
				for (let th of this.paneContent.find('thead').findAll('th')) {
					if (th.dataset.name == td.dataset.name) {
						th.find('.crater-content-options').css({ visibility: 'hidden' });
					}
				}
			});
		};

		this.paneContent.find('tbody').findAll('td').forEach(td => {
			tableBodyDataHandler(td);
		});

		let getName = () => {
			let otherThs = this.paneContent.findAll('th');
			let copyName = dataName;
			let found = true;

			while (found) {
				copyName = copyName + '-copy';
				found = false;
				for (let th of otherThs) {
					if (th.dataset.name == copyName) {
						found = true;
						break;
					}
				}
			}

			return copyName;
		};

		this.paneContent.find('thead').addEventListener('click', event => {
			let target = event.target;
			if (target.classList.contains('delete-crater-table-content-column')) {
				if (!confirm("Do you really want to delete this column")) {
					return;
				}

				let th = target.getParents('TH');

				let name = th.dataset.name;

				//remove the tds
				this.paneContent.findAll('td').forEach((td) => {
					if (td.nodeName == 'TD' && td.dataset.name == name) {
						td.remove();
					}
				});

				table.findAll('td').forEach((td) => {
					if (td.nodeName == 'TD' && td.dataset.name == name) {
						td.remove();
					}
				});

				//remove the TH
				table.find('thead').findAll('th').forEach(aTH => {
					if (aTH.dataset.name == name) {
						aTH.remove();
					}
				});
				th.remove();
			}
			else if (target.classList.contains('add-before-crater-table-content-column')) {
				let th = target.getParents('TH');
				let copyName = getName();
				//remove the tds
				this.paneContent.findAll('td').forEach((td) => {
					if (td.nodeName == 'TD' && td.dataset.name == th.dataset.name) {
						let aTDClone = td.cloneNode(true);
						aTDClone.dataset.name = copyName;
						td.before(aTDClone);
						tableBodyDataHandler(aTDClone);
					}
				});

				table.findAll('td').forEach((td) => {
					if (td.nodeName == 'TD' && td.dataset.name == th.dataset.name) {
						let aTDClone = td.cloneNode(true);
						aTDClone.dataset.name = copyName;
						td.before(aTDClone);
					}
				});

				//remove the TH
				let newPaneTH: any;
				table.find('thead').findAll('th').forEach(aTH => {
					if (aTH.dataset.name == th.dataset.name) {
						let aTHclone = aTH.cloneNode(true);
						aTHclone.dataset.name = copyName;
						aTHclone.innerText = `SAMPLE${copyName.slice(dataName.length)}`;
						aTH.before(aTHclone);
						newPaneTH = aTHclone;
					}
				});

				let aTHPaneClone = th.cloneNode(true);
				aTHPaneClone.dataset.name = copyName;
				aTHPaneClone.find('input').setAttribute('value', `${'SAMPLE'}${copyName.slice(dataName.length)}`);
				th.before(aTHPaneClone);
				tableHeadHandler(aTHPaneClone, newPaneTH);
			}
			else if (target.classList.contains('add-after-crater-table-content-column')) {
				let th = target.getParents('TH');
				let copyName = getName();
				//remove the tds
				this.paneContent.findAll('td').forEach((td) => {
					if (td.nodeName == 'TD' && td.dataset.name == th.dataset.name) {
						let aTDClone = td.cloneNode(true);
						aTDClone.dataset.name = copyName;
						td.after(aTDClone);
						tableBodyDataHandler(aTDClone);
					}
				});

				table.findAll('td').forEach((td) => {
					if (td.nodeName == 'TD' && td.dataset.name == th.dataset.name) {
						let aTDClone = td.cloneNode(true);
						aTDClone.dataset.name = copyName;
						td.after(aTDClone);
					}
				});

				//remove the TH
				let newPaneTH: any;
				table.find('thead').findAll('th').forEach(aTH => {
					if (aTH.dataset.name == th.dataset.name) {
						let aTHclone = aTH.cloneNode(true);
						aTHclone.dataset.name = copyName;
						aTHclone.innerText = `SAMPLE${copyName.slice(dataName.length)}`;
						aTH.after(aTHclone);
						newPaneTH = aTHclone;
					}
				});

				let aTHPaneClone = th.cloneNode(true);
				aTHPaneClone.dataset.name = copyName;
				aTHPaneClone.find('input').setAttribute('value', `${'SAMPLE'}${copyName.slice(dataName.length)}`);
				th.after(aTHPaneClone);
				tableHeadHandler(aTHPaneClone, newPaneTH);
			}
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());

			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart            
		});

		let tableSettings = this.paneContent.find('.table-settings');
		let tableHeaderSettings = this.paneContent.find('.table-header-settings');
		let tableBodyDataSettings = this.paneContent.find('.table-data-settings');

		tableHeaderSettings.find('#fontsize-cell').onChanged(value => {
			table.findAll('th').forEach(th => {
				th.css({ fontSize: value });
			});
		});

		tableHeaderSettings.find('#show-cell').onChanged(value => {
			if (value == 'No') {
				table.find('thead').hide();
			} else {
				table.find('thead').show();
			}
		});

		let headerColorCell = tableHeaderSettings.find('#color-cell').parentNode;
		this.pickColor({ parent: headerColorCell, cell: headerColorCell.find('#color-cell') }, (color) => {
			table.findAll('th').forEach(th => {
				th.css({ color });
			});
			headerColorCell.find('#color-cell').value = color;
			headerColorCell.find('#color-cell').setAttribute('value', color);
		});

		let headerBackgroundColorCell = tableHeaderSettings.find('#backgroundcolor-cell').parentNode;
		this.pickColor({ parent: headerBackgroundColorCell, cell: headerBackgroundColorCell.find('#backgroundcolor-cell') }, (backgroundColor) => {
			table.find('thead').find('tr').css({ backgroundColor });
			headerBackgroundColorCell.find('#backgroundcolor-cell').value = backgroundColor;
			headerBackgroundColorCell.find('#backgroundcolor-cell').setAttribute('value', backgroundColor);
		});

		tableHeaderSettings.find('#height-cell').onChanged(value => {
			table.find('thead').find('tr').css({ height: value });
		});

		tableBodyDataSettings.find('#fontsize-cell').onChanged(value => {
			table.findAll('td').forEach(th => {
				th.css({ fontSize: value });
			});
		});

		let dataColorCell = tableBodyDataSettings.find('#color-cell').parentNode;
		this.pickColor({ parent: dataColorCell, cell: dataColorCell.find('#color-cell') }, (color) => {
			table.findAll('td').forEach(td => {
				td.css({ color });
			});
			dataColorCell.find('#color-cell').value = color;
			dataColorCell.find('#color-cell').setAttribute('value', color);
		});

		let dataBackgroundColorCell = tableBodyDataSettings.find('#backgroundcolor-cell').parentNode;
		this.pickColor({ parent: dataBackgroundColorCell, cell: dataBackgroundColorCell.find('#backgroundcolor-cell') }, (backgroundColor) => {
			table.find('tbody').findAll('tr').forEach(tr => {
				tr.css({ backgroundColor });
			});
			dataBackgroundColorCell.find('#backgroundcolor-cell').value = backgroundColor;
			dataBackgroundColorCell.find('#backgroundcolor-cell').setAttribute('value', backgroundColor);
		});

		tableBodyDataSettings.find('#backgroundcolor-cell').onChanged(value => {
			table.find('tbody').findAll('tr').forEach(tr => {
				tr.css({ backgroundColor: value });
			});
		});

		tableBodyDataSettings.find('#height-cell').onChanged(value => {
			table.find('tbody').findAll('tr').forEach(tr => {
				tr.css({ height: value });
			});
		});

		tableSettings.find('#bordersize-cell').onChanged(value => {
			table.findAll('tr').forEach(tr => {
				tr.css({ borderWidth: value });
			});
		});

		let borderColorCell = tableSettings.find('#bordercolor-cell').parentNode;
		this.pickColor({ parent: borderColorCell, cell: borderColorCell.find('#bordercolor-cell') }, (borderColor) => {
			table.findAll('tr').forEach(tr => {
				tr.css({ borderColor });
			});
			borderColorCell.find('#bordercolor-cell').value = borderColor;
			borderColorCell.find('#bordercolor-cell').setAttribute('value', borderColor);
		});

		tableSettings.find('#borderstyle-cell').onChanged(value => {
			table.findAll('tr').forEach(tr => {
				tr.css({ borderStyle: value });
			});
		});
	}

	public update(params) {
		this.element = params.element;
		this.key = this.element.dataset.key;
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;
		this.paneContent = this.setUpPaneContent(params);

		let headers = this.sharePoint.properties.pane.content[this.key].settings.headers.toString();

		let paneConnection = this.sharePoint.app.find('.crater-property-connection');

		let metaWindow = this.elementModifier.createForm({
			title: 'Set Table Sample', attributes: { id: 'meta-form', class: 'form' },
			contents: {
				Names: { element: 'input', attributes: { id: 'meta-data-names', name: 'Names', value: headers }, options: params.options, note: 'Names of data should be comma seperated[data1, data2]' },
			},
			buttons: {
				submit: { element: 'button', attributes: { id: 'set-meta', class: 'btn' }, text: 'Set' },
			}
		});

		metaWindow.find('#set-meta').addEventListener('click', event => {
			event.preventDefault();

			let names = metaWindow.find('#meta-data-names').value.split(',');
			let contents = {};

			for (let i in names) {
				names[i] = func.trem(names[i]);
				contents[names[i]] = { element: 'select', attributes: { id: 'meta-data-' + names[i], name: func.capitalize(names[i]) }, options: params.options };
			}

			this.sharePoint.properties.pane.content[this.key].settings.headers = names;

			let updateWindow = this.elementModifier.createForm({
				title: 'Setup Meta Data', attributes: { id: 'meta-data-form', class: 'form' },
				contents,
				buttons: {
					submit: { element: 'button', attributes: { id: 'update-element', class: 'btn' }, text: 'Update' },
				}
			});

			let data: any = {};
			let source: any;

			updateWindow.find('#update-element').addEventListener('click', updateEvent => {
				event.preventDefault();
				let formData = updateWindow.findAll('.form-data');

				for (let i = 0; i < formData.length; i++) {
					data[formData[i].name.toLowerCase()] = formData[i].value;
				}

				source = func.extractFromJsonArray(data, params.source);

				let newContent = this.render({ source });
				draftDom.find('.crater-table').innerHTML = newContent.find('.crater-table').innerHTML;

				this.sharePoint.properties.pane.content[this.key].draft.html = draftDom.outerHTML;

				this.paneContent.find('.table-pane').innerHTML = this.generatePaneContent({ table: newContent.find('table') }).innerHTML;

				this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			});

			let parent = metaWindow.parentNode;
			parent.innerHTML = '';
			parent.append(metaWindow, updateWindow);
		});

		if (!func.isnull(paneConnection)) {
			paneConnection.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {

				this.element.innerHTML = draftDom.innerHTML;

				this.element.css(draftDom.css());

				this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart
			});
		}

		return metaWindow;
	}
}

class Panel extends CraterWebParts {
	public params: any;
	public elementModifier = new ElementModifier();
	private element: any;
	private paneContent: any;
	private key: any;

	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.params = params;
	}

	public render() {
		let panel = this.createKeyedElement({
			element: 'div', attributes: { class: 'crater-panel crater-component crater-container', 'data-type': 'panel' }, options: ['Append', 'Edit', 'Delete', 'Clone'], children: [
				{
					element: 'div', attributes: { class: 'crater-panel-title' }, children: [
						{ element: 'p', attributes: { class: 'crater-panel-title-text' }, text: 'Panel Title' },
						{ element: 'a', attributes: { class: 'crater-panel-title-link btn' }, text: 'Link' }
					]
				},
				{ element: 'div', attributes: { class: 'crater-panel-content' } }
			]
		});
		return panel;
	}

	public rendered(params) {
		this.element = params.element;
		this.key = this.element.dataset.key;

		this.showOptions(this.element);
	}

	private setUpPaneContent(params) {
		this.element = params.element;
		this.paneContent = this.elementModifier.createElement({
			element: 'div',
			attributes: { class: 'crater-property-content' }
		});
		this.key = this.element.dataset.key;

		let view = this.sharePoint.properties.pane.content[this.key].settings.view;

		if (this.sharePoint.properties.pane.content[this.key].draft.pane.content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.pane.content;
		}
		else if (this.sharePoint.properties.pane.content[this.key].content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].content;
		}
		else {
			this.paneContent.makeElement({
				element: 'div', children: [
					this.elementModifier.createElement({
						element: 'button', attributes: { class: 'btn new-component', style: { display: 'inline-block', borderRadius: '5px' } }, text: 'Add New'
					})
				]
			});

			this.paneContent.append(this.createKeyedElement({
				element: 'div', attributes: { class: 'title-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Panel Title'
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'title', value: this.element.find('.crater-panel-title').textContent
							}),
							this.elementModifier.cell({
								element: 'input', name: 'backgroundcolor', value: this.element.find('.crater-panel-title').css()['background-color'], list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'color', value: this.element.find('.crater-panel-title').css().color, list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontsize', value: this.element.find('.crater-panel-title').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'height', value: this.element.find('.crater-panel-title').css()['height']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'width', value: this.element.find('.crater-panel-title').css()['width']
							}),
							this.elementModifier.cell({
								element: 'select', name: 'layout', options: ['Full', 'Left', 'Center', 'Right']
							}),
							this.elementModifier.cell({
								element: 'select', name: 'position', options: ['Left', 'Center', 'Right']
							}),
						]
					})
				]
			}));

			this.paneContent.append(this.createKeyedElement({
				element: 'div', attributes: { class: 'title-link-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Title Link Settings'
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'color', list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'background color', list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'url'
							}),
							this.elementModifier.cell({
								element: 'input', name: 'text'
							}),
							this.elementModifier.cell({
								element: 'input', name: 'border', list: func.borders
							}),
							this.elementModifier.cell({
								element: 'select', name: 'Show', options: ['Yes', 'No']
							})
						]
					})
				]
			}));

			let sectionContentsPane = this.paneContent.makeElement({
				element: 'div', attributes: { class: 'card panel-contents-pane' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Panel Contents'
							})
						]
					}),
				]
			});

			let contents = this.element.find('.crater-panel-content').findAll('.keyed-element');

			for (let i = 0; i < contents.length; i++) {
				sectionContentsPane.makeElement({
					element: 'div',
					attributes: {
						style: { border: '1px solid #bbbbbb', margin: '.5em 0em' }, class: 'crater-panel-content-row-pane row'
					},
					children: [
						this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-panel-content-row' }),
						this.elementModifier.createElement({
							element: 'h3', attributes: { id: 'name' }, text: contents[i].dataset.type.toUpperCase()
						})
					]
				});
			}

			this.paneContent.append(this.createKeyedElement({
				element: 'div', attributes: { class: 'settings-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Panel Settings'
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'select', name: 'Box Content', options: ['Yes', 'No']
							})
						]
					})
				]
			}));
		}

		return this.paneContent;
	}

	private listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		this.paneContent = this.sharePoint.app.find('.crater-property-content').monitor();
		let panelContents = this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-panel-content');
		let panelContentDom = panelContents.childNodes;
		let panelContentPane = this.paneContent.find('.panel-contents-pane');

		let view = this.sharePoint.properties.pane.content[this.key].settings.view;
		let panelContentPanePrototype = this.elementModifier.createElement({
			element: 'div',
			attributes: {
				style: { border: '1px solid #bbbbbb', margin: '.5em 0em' }, class: 'crater-panel-content-row-pane row'
			},
			children: [
				this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-panel-content-row' }),
				this.elementModifier.createElement({
					element: 'h3', attributes: { id: 'name' }, text: 'New Webpart'
				})
			]
		});

		let panelContentRowHandler = (panelContentRowPane, panelContentRowDom) => {

			panelContentRowPane.find('#name').innerHTML = panelContentRowDom.dataset.type.toUpperCase();

			panelContentRowPane.addEventListener('mouseover', event => {
				panelContentRowPane.find('.crater-content-options').css({ visibility: 'visible' });
			});

			panelContentRowPane.addEventListener('mouseout', event => {
				panelContentRowPane.find('.crater-content-options').css({ visibility: 'hidden' });
			});

			panelContentRowPane.find('.delete-crater-panel-content-row').addEventListener('click', event => {
				panelContentRowDom.remove();
				panelContentRowPane.remove();
			});

			panelContentRowPane.find('.add-before-crater-panel-content-row').addEventListener('click', event => {
				this.paneContent.append(
					this.sharePoint.displayPanel(webpart => {
						let newPanelContent = this.sharePoint.appendWebpart(panelContents, webpart.dataset.webpart);
						panelContentRowDom.before(newPanelContent.cloneNode(true));
						newPanelContent.remove();

						let newSectionContentRow = panelContentPanePrototype.cloneNode(true);
						panelContentRowPane.after(newSectionContentRow);

						panelContentRowHandler(newSectionContentRow, newPanelContent);
					})
				);
			});

			panelContentRowPane.find('.add-after-crater-panel-content-row').addEventListener('click', event => {
				this.paneContent.append(
					this.sharePoint.displayPanel(webpart => {
						let newPanelContent = this.sharePoint.appendWebpart(panelContents, webpart.dataset.webpart);
						panelContentRowDom.after(newPanelContent.cloneNode(true));
						newPanelContent.remove();

						let newPanelContentRow = panelContentPanePrototype.cloneNode(true);
						panelContentRowPane.after(newPanelContentRow);

						panelContentRowHandler(newPanelContentRow, newPanelContent);
					})
				);
			});
		};

		let titlePane = this.paneContent.find('.title-pane');
		let titleLinkPane = this.paneContent.find('.title-link-pane');
		let settingsPane = this.paneContent.find('.settings-pane');

		let title = this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-panel-title');
		titlePane.find('#height-cell').onChanged(height => {
			title.css({ height });
		});

		titlePane.find('#width-cell').onChanged(width => {
			title.css({ width });
		});

		titlePane.find('#position-cell').onChanged(position => {
			if (position == 'Center') {
				title.css({ alignSelf: 'Center' });
			}
			else if (position == 'Right') {
				title.css({ alignSelf: 'flex-end' });
			}
			else {
				title.css({ alignSelf: 'flex-start' });
			}
		});

		titlePane.find('#layout-cell').onChanged(layout => {
			if (layout == 'Left') {
				title.css({ justifyContent: 'flex-start' });
			}
			else if (layout == 'Right') {
				title.css({ justifyContent: 'flex-end' });
			}
			else if (layout == 'Center') {
				title.css({ justifyContent: 'center' });
			}
			else {
				title.css({ justifyContent: 'space-around' });
			}
		});

		titlePane.find('#title-cell').onChanged(value => {
			title.find('.crater-panel-title-text').innerText = value;
		});

		titleLinkPane.find('#text-cell').onChanged(value => {
			title.find('.crater-panel-title-link').innerText = value;
		});

		titleLinkPane.find('#color-cell').onChanged(color => {
			title.find('.crater-panel-title-link').css({ color });
		});

		titleLinkPane.find('#background-color-cell').onChanged(backgroundColor => {
			title.find('.crater-panel-title-link').css({ backgroundColor });
		});

		titleLinkPane.find('#border-cell').onChanged(border => {
			title.find('.crater-panel-title-link').css({ border });
		});

		titleLinkPane.find('#url-cell').onChanged(value => {
			title.find('.crater-panel-title-link').href = value;
		});

		titleLinkPane.find('#Show-cell').onChanged(value => {
			if (value == 'No') {
				title.find('.crater-panel-title-link').css({ display: 'none' });
			}
			else {
				title.find('.crater-panel-title-link').cssRemove(['display']);
			}
		});

		let backgroundColorCell = titlePane.find('#backgroundcolor-cell').parentNode;
		this.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.find('#backgroundcolor-cell') }, (backgroundColor) => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-panel-title').css({ backgroundColor });
			backgroundColorCell.find('#backgroundcolor-cell').value = backgroundColor;
			backgroundColorCell.find('#backgroundcolor-cell').setAttribute('value', backgroundColor);
			this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-panel-content').css({
				borderColor: backgroundColor
			});
		});

		let colorCell = titlePane.find('#color-cell').parentNode;
		this.pickColor({ parent: colorCell, cell: colorCell.find('#color-cell') }, (color) => {
			title.find('.crater-panel-title-text').css({ color });
			colorCell.find('#color-cell').value = color;
			colorCell.find('#color-cell').setAttribute('value', color);
		});

		titlePane.find('#fontsize-cell').onChanged(value => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-panel-title').css({ fontSize: value });
		});

		settingsPane.find('#Box-Content-cell').onChanged(value => {
			if (value == 'Yes') {
				this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-panel-content').css({
					borderColor: title.css().backgroundColor
				});
			} else {
				this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-panel-content').css({
					borderColor: 'transparent'
				});
			}
		});

		this.paneContent.find('.new-component').addEventListener('click', event => {

			this.sharePoint.app.findAll('.crater-display-panel').forEach(panel => {
				panel.remove();
			});

			this.paneContent.append(this.sharePoint.displayPanel(webpart => {
				let newPanelContent = this.sharePoint.appendWebpart(panelContents, webpart.dataset.webpart);
				let newPanelContentRow = panelContentPanePrototype.cloneNode(true);
				panelContentPane.append(newPanelContentRow);

				panelContentRowHandler(newPanelContentRow, newPanelContent);
			}));
		});

		this.paneContent.findAll('.crater-panel-content-row-pane').forEach((panelContent, position) => {
			panelContentRowHandler(panelContent, panelContentDom[position]);
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;

			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());

			for (let keyedElement of this.element.findAll('.keyed-element')) {
				this[keyedElement.dataset.type]({ action: 'rendered', element: keyedElement, sharePoint: this.sharePoint });
			}
		});
	}
}

class CountDown extends CraterWebParts {
	private params: any;
	public element: any;
	public key: any;
	private paneContent: any;
	private interval: any;

	constructor(params) {
		super(params);
	}

	public render(params) {

		let endDate = Math.floor(func.dateValue(func.today())) + 2;

		let endTime = 60 * 60 * 6;

		let date = {
			days: 69,
			hours: 5,
			minutes: 21,
			seconds: 3
		};

		let countDown = this.createKeyedElement({
			element: 'div', attributes: { class: 'crater-countdown crater-component', 'data-type': 'countdown' }, children: [
				{
					element: 'div', attributes: { class: 'crater-countdown-content' }, children: [
						{
							element: 'span', attributes: { class: 'crater-countdown-block crater-countdown-days' }, children: [
								{ element: 'span', attributes: { class: 'crater-countdown-counting' }, text: date.days },
								{ element: 'h2', attributes: { class: 'crater-countdown-label' }, text: 'Days' }
							]
						},
						{
							element: 'span', attributes: { class: 'crater-countdown-block crater-countdown-hours' }, children: [
								{ element: 'span', attributes: { class: 'crater-countdown-counting' }, text: date.hours },
								{ element: 'h2', attributes: { class: 'crater-countdown-label' }, text: 'Hours' }
							]
						},
						{
							element: 'span', attributes: { class: 'crater-countdown-block crater-countdown-minutes' }, children: [
								{ element: 'span', attributes: { class: 'crater-countdown-counting' }, text: date.minutes },
								{ element: 'h2', attributes: { class: 'crater-countdown-label' }, text: 'Minutes' }
							]
						},
						{
							element: 'span', attributes: { class: 'crater-countdown-block crater-countdown-seconds' }, children: [
								{ element: 'span', attributes: { class: 'crater-countdown-counting' }, text: date.seconds },
								{ element: 'h2', attributes: { class: 'crater-countdown-label' }, text: 'Seconds' }
							]
						},
					]
				}
			]
		});

		countDown.dataset.date = endDate;
		countDown.dataset.time = endTime;
		this.key = countDown.dataset.key;

		this.sharePoint.properties.pane.content[this.key].settings.date = endDate;
		this.sharePoint.properties.pane.content[this.key].settings.time = endTime;

		return countDown;
	}

	public rendered(params) {
		this.params = params;
		this.element = params.element;
		this.key = this.element.dataset.key;

		let date: any;

		this.element.dataset.date = this.sharePoint.properties.pane.content[this.key].settings.date;

		this.element.dataset.time = this.sharePoint.properties.pane.content[this.key].settings.time;

		let secondsInCurrentDaysBefore = func.secondsInDays(Math.floor(this.sharePoint.properties.pane.content[this.key].settings.date));

		let secondsInCurrentTime = Math.floor(this.sharePoint.properties.pane.content[this.key].settings.time);

		let secondsTogoCurrently = secondsInCurrentDaysBefore + secondsInCurrentTime;

		clearInterval(this.sharePoint.properties.pane.content[this.key].settings.interval);

		date = this.getDate(secondsTogoCurrently);

		this.element.find('.crater-countdown-days').find('.crater-countdown-counting').innerText = date.days;

		this.element.find('.crater-countdown-hours').find('.crater-countdown-counting').innerText = date.hours;

		this.element.find('.crater-countdown-minutes').find('.crater-countdown-counting').innerText = date.minutes;

		this.element.find('.crater-countdown-seconds').find('.crater-countdown-counting').innerText = date.seconds;

		this.sharePoint.properties.pane.content[this.key].settings.interval = setInterval(() => {
			date = this.getDate(secondsTogoCurrently);

			if (date.past) {
				this.element.classList.toggle('crater-countdown-past');
			} else {
				this.element.classList.remove('crater-countdown-past');
			}

			this.element.find('.crater-countdown-days').find('.crater-countdown-counting').innerText = date.days;

			this.element.find('.crater-countdown-hours').find('.crater-countdown-counting').innerText = date.hours;

			this.element.find('.crater-countdown-minutes').find('.crater-countdown-counting').innerText = date.minutes;

			this.element.find('.crater-countdown-seconds').find('.crater-countdown-counting').innerText = date.seconds;
		}, 1000);
	}

	public setUpPaneContent(params) {
		this.element = params.element;
		this.paneContent = this.elementModifier.createElement({
			element: 'div',
			attributes: { class: 'crater-property-content' }
		});
		this.key = this.element.dataset.key;

		if (this.sharePoint.properties.pane.content[this.key].draft.pane.content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.pane.content;
		}
		else if (this.sharePoint.properties.pane.content[this.key].content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].content;
		}
		else {
			this.paneContent.makeElement({ element: 'div' });

			let countingPane = this.paneContent.makeElement({
				element: 'div', attributes: { class: 'counting-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Countdown Counts'
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'Color', list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'BackgroundColor', list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'FontSize',
							}),
							this.elementModifier.cell({
								element: 'input', name: 'FontStyle',
							}),
						]
					})
				]
			});

			let labelPane = this.paneContent.makeElement({
				element: 'div', attributes: { class: 'label-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Countdown Label'
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'Color', list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'BackgroundColor', list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'FontSize',
							}),
							this.elementModifier.cell({
								element: 'input', name: 'FontStyle',
							}),
						]
					})
				]
			});

			let settingsPane = this.paneContent.makeElement({
				element: 'div', attributes: { class: 'settings-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Countdown Settings'
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'Date', dataAttributes: { type: 'date' }
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Time', dataAttributes: { type: 'time' }
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Border',
							}),
							this.elementModifier.cell({
								element: 'input', name: 'BorderRadius',
							}),
						]
					})
				]
			});
		}

		return this.paneContent;
	}

	public listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		this.paneContent = this.sharePoint.app.find('.crater-property-content').monitor();

		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;

		let countingPane = this.paneContent.find('.counting-pane');

		let countingColorCell = countingPane.find('#Color-cell').parentNode;
		this.pickColor({ parent: countingColorCell, cell: countingColorCell.find('#Color-cell') }, (color) => {
			draftDom.findAll('.crater-countdown-counting').forEach(element => {
				element.css({ color });
			});
			countingColorCell.find('#Color-cell').value = color;
			countingColorCell.find('#Color-cell').setAttribute('value', color);
		});

		let countingBackgroundColorCell = countingPane.find('#BackgroundColor-cell').parentNode;
		this.pickColor({ parent: countingBackgroundColorCell, cell: countingBackgroundColorCell.find('#BackgroundColor-cell') }, (backgroundColor) => {
			draftDom.findAll('.crater-countdown-counting').forEach(element => {
				element.css({ backgroundColor });
			});
			countingBackgroundColorCell.find('#BackgroundColor-cell').value = backgroundColor;
			countingBackgroundColorCell.find('#BackgroundColor-cell').setAttribute('value', backgroundColor);
		});

		countingPane.find('#FontSize-cell').onChanged(fontSize => {
			draftDom.findAll('.crater-countdown-counting').forEach(element => {
				element.css({ fontSize });
			});
		});

		countingPane.find('#FontStyle-cell').onChanged(fontFamily => {
			draftDom.findAll('.crater-countdown-counting').forEach(element => {
				element.css({ fontFamily });
			});
		});

		let labelPane = this.paneContent.find('.label-pane');

		let labelColorCell = labelPane.find('#Color-cell').parentNode;
		this.pickColor({ parent: labelColorCell, cell: labelColorCell.find('#Color-cell') }, (color) => {
			draftDom.findAll('.crater-countdown-label').forEach(element => {
				element.css({ color });
			});
			labelColorCell.find('#Color-cell').value = color;
			labelColorCell.find('#Color-cell').setAttribute('value', color);
		});

		let labelBackgroundColorCell = labelPane.find('#BackgroundColor-cell').parentNode;
		this.pickColor({ parent: labelBackgroundColorCell, cell: labelBackgroundColorCell.find('#BackgroundColor-cell') }, (backgroundColor) => {
			draftDom.findAll('.crater-countdown-label').forEach(element => {
				element.css({ backgroundColor });
			});
			labelBackgroundColorCell.find('#BackgroundColor-cell').value = backgroundColor;
			labelBackgroundColorCell.find('#BackgroundColor-cell').setAttribute('value', backgroundColor);
		});

		labelPane.find('#FontSize-cell').onChanged(size => {
			draftDom.findAll('.crater-countdown-label').forEach(element => {
				element.css({ fontSize: size });
			});
		});

		labelPane.find('#FontStyle-cell').onChanged(style => {
			draftDom.findAll('.crater-countdown-label').forEach(element => {
				element.css({ fontFamily: style });
			});
		});

		let settingsPane = this.paneContent.find('.settings-pane');

		let settingsDate = settingsPane.find('#Date-cell');
		let settingsTime = settingsPane.find('#Time-cell');
		let settingsBorder = settingsPane.find('#Border-cell');
		let settingsBorderRadius = settingsPane.find('#BorderRadius-cell');

		settingsDate.onChanged(date => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.dataset.date = func.dateValue(date);
		});

		settingsTime.onChanged(time => {
			if (func.isTimeValid(time)) {
				this.sharePoint.properties.pane.content[this.key].draft.dom.dataset.time = func.isTimeValid(time);
			}
		});

		settingsBorder.onChanged(border => {
			draftDom.findAll('.crater-countdown-block').forEach(element => {
				element.css({ border });
			});
		});

		settingsBorderRadius.onChanged(borderRadius => {
			draftDom.findAll('.crater-countdown-block').forEach(element => {
				element.css({ borderRadius });
			});
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;

			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());

			let secondsInDayBefore = func.secondsInDays(Math.floor(this.sharePoint.properties.pane.content[this.key].draft.dom.dataset.date));

			let secondsInTime = Math.floor(this.sharePoint.properties.pane.content[this.key].draft.dom.dataset.time);

			let secondsInCurrentDaysBefore = func.secondsInDays(Math.floor(this.sharePoint.properties.pane.content[this.key].settings.date));

			let secondsInCurrentTime = Math.floor(this.sharePoint.properties.pane.content[this.key].settings.time);

			let secondsTogo = secondsInDayBefore + secondsInTime;

			let secondsTogoCurrently = secondsInCurrentDaysBefore + secondsInCurrentTime;

			if (secondsTogo != secondsTogoCurrently) {
				this.sharePoint.properties.pane.content[this.key].settings.date = this.sharePoint.properties.pane.content[this.key].draft.dom.dataset.date;

				this.sharePoint.properties.pane.content[this.key].settings.time = this.sharePoint.properties.pane.content[this.key].draft.dom.dataset.time;
			}
		});
	}

	private getDate(endTime) {
		let date = func.secondsInDays(func.dateValue(func.today()));
		let time = func.timeToday();
		let present = (endTime - (date + time) >= 0);

		let dateObject: any = func.getDateObject(
			present
				? endTime - (date + time)
				: (date + time) - endTime
		);
		if (present) dateObject.past = false;
		else dateObject.past = true;

		return dateObject;
	}
}

class DateList extends CraterWebParts {
	public params;
	public element;
	public key;
	public paneContent;
	public elementModifier = new ElementModifier();
	public monthArray = func.trimMonthArray();


	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.params = params;
	}

	public render(params) {

		if (!func.isset(params.source)) {
			params.source = [
				{
					day: "19",
					month: "Aug",
					title: "DateList Item 1",
					subtitle: "Lagos Island, Lagos",
					body: 'Donec ut maximus magna. Quisque id placerat ex. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Cras quis tellus quis orci tempus feugiat ac eu orci.'
				},
				{
					day: "15",
					month: "jul",
					title: "DateList Item 2",
					subtitle: "Lagos Island, Lagos",
					body: 'Donec ut maximus magna. Quisque id placerat ex. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Cras quis tellus quis orci tempus feugiat ac eu orci.'
				},
				{
					day: "28",
					month: "oct",
					title: "DateList Item 3",
					subtitle: "Lagos Island, Lagos",
					body: 'Donec ut maximus magna. Quisque id placerat ex. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Cras quis tellus quis orci tempus feugiat ac eu orci.'
				}
			];
		}

		let dateList = this.createKeyedElement({
			element: 'div', attributes: { class: 'crater-datelist', 'data-type': 'datelist' }, children: [
				{
					element: 'div', attributes: { class: 'crater-datelist-title' }, children: [
						{ element: 'img', attributes: { class: 'crater-datelist-title-imgIcon', src: this.sharePoint.images.async } },
						{ element: 'span', attributes: { class: 'crater-datelist-title-captionTitle' }, text: 'Date-List' }
					]
				},
				{ element: 'div', attributes: { class: 'crater-datelist-content' } }
			]
		});

		let dateListContent = dateList.find(`.crater-datelist-content`);


		for (let row of params.source) {
			dateListContent.makeElement(
				{
					element: 'div', attributes: { class: 'crater-datelist-content-item' }, children: [
						this.elementModifier.createElement({
							element: 'div', attributes: { class: 'crater-datelist-content-item-date' }, children: [
								{
									element: 'div', attributes: { class: 'crater-datelist-content-item-date-day', id: 'Day' }, text: row.day
								},
								{ element: 'div', attributes: { class: 'crater-datelist-content-item-date-month', id: 'Month' }, text: row.month.toUpperCase() }
							]
						}),
						this.elementModifier.createElement({
							element: 'div', attributes: { class: 'crater-datelist-content-item-text' }, children: [
								{ element: 'div', attributes: { class: 'crater-datelist-content-item-text-main', id: 'mainText' }, text: row.title },
								{ element: 'div', attributes: { class: 'crater-datelist-content-item-text-subtitle', id: 'subtitle' }, text: row.subtitle },
								{ element: 'div', attributes: { class: 'crater-datelist-content-item-text-body', id: 'body' }, text: row.body },
							]
						})
					]
				}
			);
		}

		this.key = this.key || dateList.dataset.key;
		return dateList;
	}

	public rendered(params) {
		this.element = params.element;
	}

	public setUpPaneContent(params) {
		let key = params.element.dataset['key'];//create a key variable and set it to the webpart key
		this.element = this.sharePoint.properties.pane.content[key].draft.dom;//define the declared element to the draft dom content
		this.paneContent = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'crater-property-content' }
		}).monitor(); //monitor the content pane 
		if (this.sharePoint.properties.pane.content[key].draft.pane.content.length !== 0) {//check if draft pane content is not empty and set it to the pane content
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[key].draft.pane.content;
		}
		else if (this.sharePoint.properties.pane.content[key].content.length !== 0) {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[key].content;
		} else {
			let dateList = this.sharePoint.properties.pane.content[key].draft.dom.find('.crater-datelist-content');
			let dateListRows = dateList.findAll('.crater-datelist-content-item');
			this.paneContent.makeElement({
				element: 'div', children: [
					this.elementModifier.createElement(
						{ element: 'button', attributes: { class: 'btn new-component', style: { display: 'inline-block', borderRadius: '5px' } }, text: 'Add New' }
					)
				]
			});


			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'title-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							{ element: 'h2', attributes: { class: 'title' }, text: 'Date-List Title Layout' }
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [//create the cells for changing crater event title
							this.elementModifier.cell({
								element: 'img', name: 'icon', attributes: {}, dataAttributes: { class: 'crater-icon', src: this.element.find('.crater-datelist-title-imgIcon').src }
							}),
							this.elementModifier.cell({
								element: 'input', name: 'title', value: this.element.find('.crater-datelist-title').textContent
							}),
							this.elementModifier.cell({
								element: 'input', name: 'backgroundcolor', value: this.element.find('.crater-datelist-title').css()['background-color'], list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'color', value: this.element.find('.crater-datelist-title').css().color, list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontsize', value: this.element.find('.crater-datelist-title').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'height', value: this.element.find('.crater-datelist-title').css()['height'] || ''
							}),
							this.elementModifier.cell({
								element: 'select', name: 'toggleTitle', options: ['show', 'hide']
							})
						]
					})
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'datelist-date-row-pane card' }, children: [
					{
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Edit Date-List Dates'
							}),
						]
					},
					{
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'daySize', value: this.element.find('.crater-datelist-content-item-date-day').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'monthSize', value: this.element.find('.crater-datelist-content-item-date-month').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontFamily', value: this.element.find('.crater-datelist-content-item-date').css()['font-family']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'dayColor', value: this.element.find('.crater-datelist-content-item-date-day').css()['color'], list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'monthColor', value: this.element.find('.crater-datelist-content-item-date-month').css()['color'], list: func.colors
							}),
							this.elementModifier.cell({
								element: 'select', name: 'toggleDate', options: ['show', 'hide']
							})
						]
					}

				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'datelist-title-row-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Edit Date-List Title'
							})
						]
					}),
					{
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'fontSize', value: this.element.find('.crater-datelist-content-item-text-main').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontFamily', value: this.element.find('.crater-datelist-content-item-text-main').css()['font-family']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'titleColor', value: this.element.find('.crater-datelist-content-item-text-main').css()['color'], list: func.colors
							}),
							this.elementModifier.cell({
								element: 'select', name: 'toggleTitle', options: ['show', 'hide']
							})
						]
					}
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'datelist-subtitle-row-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Edit Date-List Subtitle'
							})
						]
					}),
					{
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'fontSize', value: this.element.find('.crater-datelist-content-item-text-subtitle').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontFamily', value: this.element.find('.crater-datelist-content-item-text-subtitle').css()['font-family']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'subtitleColor', value: this.element.find('.crater-datelist-content-item-text-subtitle').css().color, list: func.colors
							}),
							this.elementModifier.cell({
								element: 'select', name: 'toggleSubtitle', options: ['show', 'hide']
							})
						]
					}
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'datelist-body-row-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Edit Date-List Body'
							})
						]
					}),
					{
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'fontSize', value: this.element.find('.crater-datelist-content-item-text-body').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontFamily', value: this.element.find('.crater-datelist-content-item-text-body').css()['font-family']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'bodyColor', value: this.element.find('.crater-datelist-content-item-text-body').css()['color'], list: func.colors
							}),
							this.elementModifier.cell({
								element: 'select', name: 'toggleBody', options: ['show', 'hide']
							})
						]
					}
				]
			});


			this.paneContent.append(this.generatePaneContent({ list: dateListRows }));
		}
		return this.paneContent;
	}

	public generatePaneContent(params) {
		let dateListPane = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'card datelist-pane' }, children: [
				this.elementModifier.createElement({
					element: 'div', attributes: { class: 'card-title' }, children: [
						this.elementModifier.createElement({
							element: 'h2', attributes: { class: 'title' }, text: 'Date-List Rows'
						})
					]
				}),
			]
		});

		for (let i = 0; i < params.list.length; i++) {
			dateListPane.makeElement({
				element: 'div',
				attributes: { class: 'crater-datelist-item-pane row' },
				children: [
					this.paneOptions({ options: ['AA', 'AB', 'D'], owner: 'crater-datelist-content-item' }),
					this.elementModifier.cell({
						element: 'input', name: 'Day', attribute: { class: 'crater-date' }, value: params.list[i].find('#Day').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Month', attribute: { class: 'crater-date' }, value: params.list[i].find('#Month').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'title', value: params.list[i].find('#mainText').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'subtitle', value: params.list[i].find('#subtitle').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'body', value: params.list[i].find('#body').textContent
					}),
				]
			});
		}

		return dateListPane;

	}

	public listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		this.paneContent = this.sharePoint.app.find('.crater-property-content').monitor();
		//get the content and all the events
		let dateList = this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-datelist-content');
		let dateListRow = dateList.findAll('.crater-datelist-content-item');

		let dateListRowPanePrototype = this.elementModifier.createElement({//create a row on the property pane
			element: 'div',
			attributes: { class: 'crater-datelist-item-pane row' },
			children: [
				this.paneOptions({ options: ['AA', 'AB', 'D'], owner: 'crater-datelist-content-item' }),
				this.elementModifier.cell({
					element: 'input', name: 'Day', attribute: { class: 'crater-date' }, value: ''
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Month', attribute: { class: 'crater-date' }, value: ''
				}),
				this.elementModifier.cell({
					element: 'input', name: 'title', value: ''
				}),
				this.elementModifier.cell({
					element: 'input', name: 'subtitle', value: ''
				}),
				this.elementModifier.cell({
					element: 'input', name: 'body', value: ''
				}),
			]
		});


		let dateListRowDomPrototype = this.createKeyedElement(
			{
				element: 'div', attributes: { class: 'crater-datelist-content-item keyed-element' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'crater-datelist-content-item-date' }, children: [
							{
								element: 'div', attributes: { class: 'crater-datelist-content-item-date-day', id: 'Day' }, text: ''
							},
							{ element: 'div', attributes: { class: 'crater-datelist-content-item-date-month', id: 'Month' }, text: '' }
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'crater-datelist-content-item-text' }, children: [
							{ element: 'div', attributes: { class: 'crater-datelist-content-item-text-main', id: 'mainText' }, text: '' },
							{ element: 'div', attributes: { class: 'crater-datelist-content-item-text-subtitle', id: 'subtitle' }, text: '' },
							{ element: 'div', attributes: { class: 'crater-datelist-content-item-text-body', id: 'body' }, text: '' },
						]
					})
				]
			}
		);

		let dateRowHandler = (dateRowPane, dateRowDom) => {
			dateRowPane.addEventListener('mouseover', event => {
				dateRowPane.find('.crater-content-options').css({ visibility: 'visible' });
			});

			dateRowPane.addEventListener('mouseout', event => {
				dateRowPane.find('.crater-content-options').css({ visibility: 'hidden' });
			});

			// get the values of the newly created row on the property - pane
			dateRowPane.find('#title-cell').onChanged(value => {
				dateRowDom.find('.crater-datelist-content-item-text-main').innerHTML = value;
			});

			dateRowPane.find('#subtitle-cell').onChanged(value => {
				dateRowDom.find('.crater-datelist-content-item-text-subtitle').innerHTML = value;
			});

			dateRowPane.find('#body-cell').onChanged(value => {
				dateRowDom.find('.crater-datelist-content-item-text-body').innerHTML = value;
			});

			dateRowPane.find('#Day-cell').onChanged(value => {
				dateRowDom.find('.crater-datelist-content-item-date-day').innerHTML = value;
			});

			dateRowPane.find('#Month-cell').onChanged(value => {
				dateRowDom.find('.crater-datelist-content-item-date-month').innerHTML = value;
			});

			dateRowPane.find('.delete-crater-datelist-content-item').addEventListener('click', event => {
				dateRowDom.remove();
				dateRowPane.remove();
			});

			dateRowPane.find('.add-before-crater-datelist-content-item').addEventListener('click', event => {
				let newdateRowDom = dateListRowDomPrototype.cloneNode(true);
				let newdateRowPane = dateListRowPanePrototype.cloneNode(true);

				dateRowDom.before(newdateRowDom);
				dateRowPane.before(newdateRowPane);
				dateRowHandler(newdateRowPane, newdateRowDom);
			});

			dateRowPane.find('.add-after-crater-datelist-content-item').addEventListener('click', event => {
				let newdateRowDom = dateListRowDomPrototype.cloneNode(true);
				let newdateRowPane = dateListRowPanePrototype.cloneNode(true);

				dateRowDom.after(newdateRowDom);
				dateRowPane.after(newdateRowPane);

				dateRowHandler(newdateRowPane, newdateRowDom);
			});
		};

		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;

		let titlePane = this.paneContent.find('.title-pane');
		let dateListDateRowPane = this.paneContent.find('.datelist-date-row-pane');
		let dateListTitleRowPane = this.paneContent.find('.datelist-title-row-pane');
		let dateListSubtitleRowPane = this.paneContent.find('.datelist-subtitle-row-pane');
		let dateListBodyRowPane = this.paneContent.find('.datelist-body-row-pane');

		let dateListTitleParent = dateListTitleRowPane.find('#titleColor-cell').parentNode;
		this.pickColor({ parent: dateListTitleParent, cell: dateListTitleParent.find('#titleColor-cell') }, (color) => {
			draftDom.findAll('.crater-datelist-content-item-text-main').forEach(element => {
				element.css({ color });//get the color of the event font in the draftDom
			});
			dateListTitleParent.find('#titleColor-cell').value = color;
			dateListTitleParent.find('#titleColor-cell').setAttribute('value', color); //set the value of the eventColor cell in the pane to the color
		});

		let dateListSubtitleParent = dateListSubtitleRowPane.find('#subtitleColor-cell').parentNode;
		this.pickColor({ parent: dateListSubtitleParent, cell: dateListSubtitleParent.find('#subtitleColor-cell') }, (color) => {
			draftDom.findAll('.crater-datelist-content-item-text-subtitle').forEach(element => {
				element.css({ color });//get the color of the event font in the draftDom
			});
			dateListSubtitleParent.find('#subtitleColor-cell').value = color;
			dateListSubtitleParent.find('#subtitleColor-cell').setAttribute('value', color); //set the value of the eventColor cell in the pane to the color
		});

		let dateListBodyParent = dateListBodyRowPane.find('#bodyColor-cell').parentNode;
		this.pickColor({ parent: dateListBodyParent, cell: dateListBodyParent.find('#bodyColor-cell') }, (color) => {
			draftDom.findAll('.crater-datelist-content-item-text-body').forEach(element => {
				element.css({ color });//get the color of the event font in the draftDom
			});
			dateListBodyParent.find('#bodyColor-cell').value = color;
			dateListBodyParent.find('#bodyColor-cell').setAttribute('value', color); //set the value of the eventColor cell in the pane to the color
		});

		let dateListDayParent = dateListDateRowPane.find('#dayColor-cell').parentNode;
		this.pickColor({ parent: dateListDayParent, cell: dateListDayParent.find('#dayColor-cell') }, (color) => {
			draftDom.findAll('.crater-datelist-content-item-date-day').forEach(element => {
				element.css({ color });//get the color of the event font in the draftDom
			});
			dateListDayParent.find('#dayColor-cell').value = color;
			dateListDayParent.find('#dayColor-cell').setAttribute('value', color); //set the value of the eventColor cell in the pane to the color
		});

		let dateListMonthParent = dateListDateRowPane.find('#monthColor-cell').parentNode;
		this.pickColor({ parent: dateListMonthParent, cell: dateListMonthParent.find('#monthColor-cell') }, (color) => {
			draftDom.findAll('.crater-datelist-content-item-date-month').forEach(element => {
				element.css({ color });//get the color of the event font in the draftDom
			});
			dateListMonthParent.find('#monthColor-cell').value = color;
			dateListMonthParent.find('#monthColor-cell').setAttribute('value', color); //set the value of the eventColor cell in the pane to the color
		});

		let iconParent = titlePane.find('#icon-cell').parentNode;
		this.uploadImage({ parent: iconParent }, (image) => {
			iconParent.find('#icon-cell').src = image.src;
			draftDom.find('.crater-datelist-title-imgIcon').src = image.src;
		});

		titlePane.find('#title-cell').onChanged(value => {
			draftDom.find('.crater-datelist-title-captionTitle').innerHTML = value;
		});

		let backgroundColorCell = titlePane.find('#backgroundcolor-cell').parentNode;
		this.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.find('#backgroundcolor-cell') }, (backgroundColor) => {
			draftDom.find('.crater-datelist-title').css({ backgroundColor });
			backgroundColorCell.find('#backgroundcolor-cell').value = backgroundColor;
			backgroundColorCell.find('#backgroundcolor-cell').setAttribute('value', backgroundColor);
		});

		let colorCell = titlePane.find('#color-cell').parentNode;
		this.pickColor({ parent: colorCell, cell: colorCell.find('#color-cell') }, (color) => {
			draftDom.find('.crater-datelist-title').css({ color });
			colorCell.find('#color-cell').value = color;
			colorCell.find('#color-cell').setAttribute('value', color);
		});


		titlePane.find('#fontsize-cell').onChanged(value => {
			draftDom.find('.crater-datelist-title').css({ fontSize: value });
		});

		titlePane.find('#height-cell').onChanged(value => {
			draftDom.find('.crater-datelist-title').css({ height: value });
		});



		dateListTitleRowPane.find('#fontSize-cell').onChanged(value => {
			draftDom.findAll('.crater-datelist-content-item-text-main').forEach(element => {
				element.css({ fontSize: value });
			});
		});
		dateListTitleRowPane.find('#fontFamily-cell').onChanged(value => {
			draftDom.findAll('.crater-datelist-content-item-text-main').forEach(element => {
				element.css({ fontFamily: value });
			});
		});

		dateListSubtitleRowPane.find('#fontSize-cell').onChanged(value => {
			draftDom.findAll('.crater-datelist-content-item-text-subtitle').forEach(element => {
				element.css({ fontSize: value });
			});
		});
		dateListSubtitleRowPane.find('#fontFamily-cell').onChanged(value => {
			draftDom.findAll('.crater-datelist-content-item-text-subtitle').forEach(element => {
				element.css({ fontFamily: value });
			});
		});

		dateListBodyRowPane.find('#fontSize-cell').onChanged(value => {
			draftDom.findAll('.crater-datelist-content-item-text-body').forEach(element => {
				element.css({ fontSize: value });
			});
		});
		dateListBodyRowPane.find('#fontFamily-cell').onChanged(value => {
			draftDom.findAll('.crater-datelist-content-item-text-body').forEach(element => {
				element.css({ fontFamily: value });
			});
		});

		dateListDateRowPane.find('#daySize-cell').onChanged(value => {
			draftDom.findAll('.crater-datelist-content-item-date-day').forEach(element => {
				element.css({ fontSize: value });
			});
		});
		dateListDateRowPane.find('#monthSize-cell').onChanged(value => {
			draftDom.findAll('.crater-datelist-content-item-date-month').forEach(element => {
				element.css({ fontSize: value });
			});
		});

		dateListDateRowPane.find('#fontFamily-cell').onChanged(value => {
			draftDom.findAll('.crater-datelist-content-item-date').forEach(element => {
				element.css({ fontFamily: value });
			});
		});

		//appends the dom and pane prototypes to the dom and pane when you click add new
		this.paneContent.find('.new-component').addEventListener('click', event => {
			let newDateRowDom = dateListRowDomPrototype.cloneNode(true);
			let newDateRowPane = dateListRowPanePrototype.cloneNode(true);

			dateList.append(newDateRowDom);//c
			this.paneContent.find('.datelist-pane').append(newDateRowPane);
			dateRowHandler(newDateRowPane, newDateRowDom);
		});

		let paneItems = this.paneContent.findAll('.crater-datelist-item-pane');
		paneItems.forEach((dateRow, position) => {
			dateRowHandler(dateRow, dateListRow[position]);
		});

		let showHeader = titlePane.find('#toggleTitle-cell');
		showHeader.addEventListener('change', e => {

			switch (showHeader.value.toLowerCase()) {
				case "hide":
					draftDom.findAll('.crater-datelist-title').forEach(element => {
						element.style.display = "none";
					});
					break;
				case "show":
					draftDom.findAll('.crater-datelist-title').forEach(element => {
						element.style.display = "grid";
					});
					break;
				default:
					draftDom.findAll('.crater-datelist-title').forEach(element => {
						element.style.display = "none";
					});
			}
		});

		let showTitle = dateListTitleRowPane.find('#toggleTitle-cell');
		showTitle.addEventListener('change', e => {

			switch (showTitle.value.toLowerCase()) {
				case "hide":
					draftDom.findAll('.crater-datelist-content-item-text-main').forEach(element => {
						element.style.display = "none";
					});
					break;
				case "show":
					draftDom.findAll('.crater-datelist-content-item-text-main').forEach(element => {
						element.style.display = "block";
					});
					break;
				default:
					draftDom.findAll('.crater-datelist-content-item-text-main').forEach(element => {
						element.style.display = "none";
					});
			}
		});

		let showSubtitle = dateListSubtitleRowPane.find('#toggleSubtitle-cell');
		showSubtitle.addEventListener('change', e => {

			switch (showSubtitle.value.toLowerCase()) {
				case "hide":
					draftDom.findAll('.crater-datelist-content-item-text-subtitle').forEach(element => {
						element.style.display = "none";
					});
					break;
				case "show":
					draftDom.findAll('.crater-datelist-content-item-text-subtitle').forEach(element => {
						element.style.display = "block";
					});
					break;
				default:
					draftDom.findAll('.crater-datelist-content-item-text-subtitle').forEach(element => {
						element.style.display = "none";
					});
			}
		});

		let showBody = dateListBodyRowPane.find('#toggleBody-cell');
		showBody.addEventListener('change', e => {

			switch (showBody.value.toLowerCase()) {
				case "hide":
					draftDom.findAll('.crater-datelist-content-item-text-body').forEach(element => {
						element.style.display = "none";
					});
					break;
				case "show":
					draftDom.findAll('.crater-datelist-content-item-text-body').forEach(element => {
						element.style.display = "block";
					});
					break;
				default:
					draftDom.findAll('.crater-datelist-content-item-text-body').forEach(element => {
						element.style.display = "none";
					});
			}
		});

		let showDate = dateListDateRowPane.find('#toggleDate-cell');
		showDate.addEventListener('change', e => {

			switch (showDate.value.toLowerCase()) {
				case "hide":
					draftDom.findAll('.crater-datelist-content-item-date').forEach(element => {
						element.style.visibility = "hidden";
					});
					break;
				case "show":
					draftDom.findAll('.crater-datelist-content-item-date').forEach(element => {
						element.style.visibility = "visible";
					});
					break;
				default:
					draftDom.findAll('.crater-datelist-content-item-date').forEach(element => {
						element.style.visibility = "hidden";
					});
			}
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = draftDom.innerHTML;//upate the webpart
			this.element.css(draftDom.css());
			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;

		});
	}

	public update(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;
		this.paneContent = this.setUpPaneContent(params);

		let paneConnection = this.sharePoint.app.find('.crater-property-connection');

		let updateWindow = this.elementModifier.createForm({
			title: 'Setup Meta Data', attributes: { id: 'meta-data-form', class: 'form' },
			contents: {
				day: { element: 'select', attributes: { id: 'meta-data-day', name: 'Day' }, options: params.options },
				month: { element: 'select', attributes: { id: 'meta-data-month', name: 'Month' }, options: params.options },
				title: { element: 'select', attributes: { id: 'meta-data-title', name: 'Title' }, options: params.options },
				subtitle: { element: 'select', attributes: { id: 'meta-data-subtitle', name: 'Subtitle' }, options: params.options },
				body: { element: 'select', attributes: { id: 'meta-data-body', name: 'Body' }, options: params.options }
			},
			buttons: {
				submit: { element: 'button', attributes: { id: 'update-element', class: 'btn' }, text: 'Update' },
			}
		});

		let data: any = {};
		let source: any;
		updateWindow.find('#update-element').addEventListener('click', event => {
			event.preventDefault();
			data.day = updateWindow.find('#meta-data-day').value;
			data.month = updateWindow.find('#meta-data-month').value;
			data.title = updateWindow.find('#meta-data-title').value;
			data.subtitle = updateWindow.find('#meta-data-subtitle').value;
			data.body = updateWindow.find('#meta-data-body').value;
			source = func.extractFromJsonArray(data, params.source);

			let newContent = this.render({ source });
			draftDom.find('.crater-datelist-content').innerHTML = newContent.find('.crater-datelist-content').innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = draftDom.outerHTML;
			this.paneContent.find('.datelist-pane').innerHTML = this.generatePaneContent({ list: newContent.findAll('.crater-datelist-content-item') }).innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
		});


		if (!func.isnull(paneConnection)) {
			paneConnection.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
				this.element.innerHTML = draftDom.innerHTML;

				this.element.css(draftDom.css());

				this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart
			});
		}

		return updateWindow;
	}
}

class Map extends CraterWebParts {
	public key;
	public element;
	public params;
	public paneContent;
	public elementModifier = new ElementModifier();
	public self = this;
	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.params = params;
	}

	public render(params?) {

		let mapDiv = this.createKeyedElement({
			element: 'div', attributes: { class: 'crater-map', 'data-type': 'map' }, children: [
				{ element: 'div', attributes: { class: 'crater-map-div', id: 'crater-map-div' } },
				{
					element: 'script', attributes: { src: 'https://maps.googleapis.com/maps/api/js?key=AIzaSyDdGAHe_9Ghatd4wZjyc3hRdirIQ1ttcv0&callback=initMap' }
				}
			]
		});

		this.key = this.key || mapDiv.dataset.key;

		this.sharePoint.properties.pane.content[this.key].settings = { myMap: { lat: -34.067, lng: 150.067, zoom: 4, markerChecked: true, color: '' } };
		window['initMap'] = this.initMap;

		this.element = mapDiv;

		return mapDiv;
	}

	public initMap = () => {
		const mapColor = this.sharePoint.properties.pane.content[this.key].settings.myMap.color;
		let mapStyles = [
			{ elementType: 'geometry.stroke', stylers: [{ color: mapColor }, { lightness: 0 }] },
			{ elementType: 'labels.text.fill', stylers: [{ color: mapColor }, { lightness: 0 }] },
			{ elementType: 'labels.text.stroke', stylers: [{ color: '#f9f9f9' }] },
			{ featureType: 'water', elementType: 'geometry.fill', stylers: [{ color: mapColor }, { lightness: 80 }] },
			{ featureType: 'water', elementType: 'labels.text.fill', stylers: [{ color: mapColor }, { lightness: 0 }] },
			{ featureType: 'water', elementType: 'labels.text.stroke', stylers: [{ color: '#f9f9f9' }] },
			{ featureType: 'road', elementType: 'geometry.fill', stylers: [{ color: mapColor }, { lightness: 80 }] },
			{ featureType: 'landscape', elementType: 'geometry.fill', stylers: [{ color: mapColor }, { lightness: 95 }] }
		];

		let newMap = {
			lat: this.sharePoint.properties.pane.content[this.key].settings.myMap.lat,
			lng: this.sharePoint.properties.pane.content[this.key].settings.myMap.lng
		};

		const styles = (mapColor !== '') ? mapStyles : '';

		// @ts-ignore
		let map = new google.maps.Map(this.element.find('#crater-map-div'), {
			center: newMap,
			zoom: this.sharePoint.properties.pane.content[this.key].settings.myMap.zoom,
			styles
		});

		if (this.sharePoint.properties.pane.content[this.key].settings.myMap.markerChecked) {
			//@ts-ignore
			let marker = new google.maps.Marker({
				position: newMap,
				map
			});
		}
	}

	public rendered(params) {
		this.element = params.element;
		this.key = params.element.dataset['key'];
		this.element.find('#crater-map-div').innerHTML = '';
		this.element.find('script').remove();
		window['initMap'] = this.initMap;
		this.element.makeElement(
			{
				element: 'script', attributes: { src: 'https://maps.googleapis.com/maps/api/js?key=AIzaSyDdGAHe_9Ghatd4wZjyc3hRdirIQ1ttcv0&callback=initMap' }
			}
		);
		console.log(this.element);
	}

	public setUpPaneContent(params) {
		let key = params.element.dataset['key'];
		this.element = this.sharePoint.properties.pane.content[key].draft.dom;
		this.paneContent = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'crater-property-content' }
		}).monitor();

		if (this.sharePoint.properties.pane.content[key].draft.pane.content.length !== 0) {//check if draft pane content is not empty and set it to the pane content
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[key].draft.pane.content;
		}
		else if (this.sharePoint.properties.pane.content[key].content.length !== 0) {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[key].content;
		} else {
			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'map-style-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Customize Map'
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							{
								element: 'div', attributes: { class: 'message-note' }, children: [
									{
										element: 'div', attributes: { class: 'message-text' }, children: [
											{ element: 'p', text: 'NOTE:' },
											{ element: 'p', attributes: { style: { color: 'green' } }, text: `Clear the color input field to reset the map color to default.` },
											{ element: 'p', attributes: { style: { color: 'green' } }, text: `The latitude/longitude should be in the format '6.6018'.` },
											{ element: 'p', attributes: { style: { color: 'green' } }, text: `The optimal zoom level values range from 0 - 20` }
										]
									}
								]
							}
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'latitude', value: ''
							}),
							this.elementModifier.cell({
								element: 'input', name: 'longitude', value: ''
							}),
							this.elementModifier.cell({
								element: 'input', name: 'zoom', value: ''
							}),
							this.elementModifier.cell({
								element: 'input', name: 'color', list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'width', value: this.element.find('#crater-map-div').css()['width'], list: func.pixelSizes
							}),
							this.elementModifier.cell({
								element: 'input', name: 'height', value: this.element.find('#crater-map-div').css()['height'], list: func.pixelSizes
							}),
							this.elementModifier.cell({
								element: 'select', name: 'marker', options: ['show', 'hide']
							})
						]
					})
				]
			});
		}
		return this.paneContent;
	}

	public generatePaneContent(params) {
	}

	public listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;

		this.paneContent = this.sharePoint.app.find('.crater-property-content').monitor();

		let mapPane = this.paneContent.find('.map-style-pane');

		mapPane.find('#latitude-cell').onChanged(value => {
			this.sharePoint.properties.pane.content[this.key].settings.myMap.lat = parseFloat(value);
		});

		let colorValue = mapPane.find('#color-cell').parentNode;
		this.pickColor({ parent: colorValue, cell: colorValue.find('#color-cell') }, (color) => {

			colorValue.find('#color-cell').value = color;
			let hexColor = ColorPicker.rgbToHex(color);
			colorValue.find('#color-cell').setAttribute('value', hexColor); //set the value of the eventColor cell in the pane to the color
			this.sharePoint.properties.pane.content[this.key].settings.myMap.color = hexColor;
		});

		mapPane.find('#longitude-cell').onChanged(value => {
			this.sharePoint.properties.pane.content[this.key].settings.myMap.lng = parseFloat(value);
		});

		mapPane.find('#zoom-cell').onChanged(value => {
			this.sharePoint.properties.pane.content[this.key].settings.myMap.zoom = parseInt(value);
		});

		let markerValue = mapPane.find('#marker-cell');
		markerValue.addEventListener('change', e => {
			let marker = markerValue.value;
			switch (marker.toLowerCase()) {
				case 'show':
					this.sharePoint.properties.pane.content[this.key].settings.myMap.markerChecked = true;
					break;
				case 'hide':
					this.sharePoint.properties.pane.content[this.key].settings.myMap.markerChecked = false;
					break;
			}
		});

		mapPane.find('#width-cell').onChanged(value => {
			draftDom.find('#crater-map-div').css({ width: value });
			this.sharePoint.properties.pane.content[this.key].settings.myMap.width = value;
		});

		mapPane.find('#height-cell').onChanged(value => {
			draftDom.find('#crater-map-div').css({ height: value });
			this.sharePoint.properties.pane.content[this.key].settings.myMap.height = value;
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.find('#crater-map-div').innerHTML = '';
			this.element.removeChild(this.element.find('script'));
			window['initMap'] = this.initMap;
			this.element.makeElement({
				element: 'script', attributes: { src: 'https://maps.googleapis.com/maps/api/js?key=AIzaSyDdGAHe_9Ghatd4wZjyc3hRdirIQ1ttcv0&callback=initMap' }
			});

			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());

			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart
		});

	}
}

class Instagram extends CraterWebParts {
	public params;
	public element;
	public key;
	public paneContent;
	public elementModifier = new ElementModifier();
	public defaultURL = 'https://www.instagram.com/p/B0Qyddphpht/';
	public endpoint;
	public displayed = false;

	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.params = params;
	}

	public render(params?) {
		let instagramDiv = this.createKeyedElement({
			element: 'div', attributes: { class: 'crater-instagram', 'data-type': 'instagram' }, children: [
				{
					element: 'div', attributes: { class: 'crater-instagram-error' }, children: [
						{ element: 'p', text: 'Are you sure that URL is valid? please enter a URL shortcode' }
					]
				},
				{ element: 'div', attributes: { class: 'crater-instagram-content' } }
			]
		});

		this.sharePoint.properties.pane.content[instagramDiv.dataset.key].settings = { myInstagram: {} };
		this.key = instagramDiv.dataset.key;
		this.element = instagramDiv;
		return instagramDiv;
	}

	public getEmbed(params?) {
		let width = this.element.getBoundingClientRect().width;
		window.onerror = (message, url, line, column, error) => {
			console.log(message, url, line, column, error);
		};
		let fetchHtml: any = new Promise((resolve, reject) => {
			let urlPoint = 'https://api.instagram.com/oembed?url=' + this.defaultURL + `&amp;maxwidth=${width}&amp;minwidth=${320}&ampomitscript=true`;
			if (func.isset(params)) {
				urlPoint = 'https://api.instagram.com/oembed?url=' + params;
			}

			let getData = fetch(urlPoint);
			resolve(getData);
			reject(new Error('couldn\'t fetch data'));
		});

		fetchHtml
			.then(response => {
				return response.json();
			})
			.then(responseData => {
				if (this.displayed) {
					let instaDiv = this.sharePoint.app.find('.crater-instagram');
					instaDiv.removeChild(instaDiv.find('.crater-instagram-content'));
					let child = this.elementModifier.createElement({
						element: 'div', attributes: { class: 'crater-instagram-content' }
					});
					instaDiv.appendChild(child);
					this.renderInstagramPost(responseData.html);

				} else {
					this.renderInstagramPost(responseData.html);
					this.displayed = true;
				}
			}).catch(error => {
				let errorMessage = this.sharePoint.app.find('.crater-instagram-error');
				errorMessage.style.display = 'block';
			});
	}

	public renderInstagramPost(params?) {
		this.element = this.sharePoint.app.find('.crater-instagram');
		let errorMessage = this.sharePoint.app.find('.crater-instagram-error');
		errorMessage.style.display = 'none';
		this.key = this.element.dataset['key'];
		let display = params;
		let instaContent = this.element.find('.crater-instagram-content');
		instaContent.innerHTML = '';

		let embedScript = this.element.createElement('script');
		embedScript.setAttribute('src', '//www.instagram.com/embed.js');
		embedScript.setAttribute('async', '');
		instaContent.appendChild(embedScript);
		instaContent.innerHTML += display;
		embedScript.addEventListener('load', e => {
			//@ts-ignore
			instgrm.Embeds.process();
		});

	}

	public rendered(params) {
		this.element = params.element;
		this.key = this.element.dataset.key;
		this.getEmbed();
	}

	public setUpPaneContent(params) {
		let key = params.element.dataset['key'];
		this.element = this.sharePoint.properties.pane.content[key].draft.dom;
		this.paneContent = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'crater-property-content' }
		});

		if (this.sharePoint.properties.pane.content[key].draft.pane.content.length !== 0) {//check if draft pane content is not empty and set it to the pane content
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[key].draft.pane.content;
		}
		else if (this.sharePoint.properties.pane.content[key].content.length !== 0) {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[key].content;
		} else {
			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'instagram-pane' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Edit Instagram Properties'
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'postUrl', value: this.defaultURL
							}),
							this.elementModifier.cell({
								element: 'select', name: 'hideCaption', options: ['show', 'hide']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'width', attributes: { placeholder: 'Please enter a width' }, value: this.element.find('.crater-instagram-content').css()['height'] || ''
							})
						]
					})
				]
			});

		}
		return this.paneContent;
	}

	public generatePaneContent(params) {

	}

	public listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;

		this.paneContent = this.sharePoint.app.find('.crater-property-content').monitor();

		let instagramPane = this.paneContent.find('.instagram-pane');
		let postUrl = instagramPane.find('#postUrl-cell');
		postUrl.onChanged(value => {
			this.sharePoint.properties.pane.content[this.key].settings.myInstagram.instaURL = value;
			this.defaultURL = this.sharePoint.properties.pane.content[this.key].settings.myInstagram.instaURL;
			let finalWidth = this.sharePoint.properties.pane.content[this.key].settings.myInstagram.instaWidth || '&amp;minwidth=320&amp;maxwidth=320';
			const finalHide = (this.sharePoint.properties.pane.content[this.key].settings.myInstagram.instaCaption) ? '&amp;hidecaption=true' : '';

			this.sharePoint.properties.pane.content[this.key].draft.newEndPoint = this.sharePoint.properties.pane.content[this.key].settings.myInstagram.instaURL + `&amp;omitscript=true${finalHide}${finalWidth}`;
		});

		let hideCaption = instagramPane.find('#hideCaption-cell');
		hideCaption.addEventListener('change', e => {
			let finalWidth = this.sharePoint.properties.pane.content[this.key].settings.myInstagram.instaWidth || '&amp;minwidth=320&amp;maxwidth=320';
			let finalURL = this.sharePoint.properties.pane.content[this.key].settings.myInstagram.instaURL || this.defaultURL;
			switch (hideCaption.value.toLowerCase()) {
				case 'hide':
					this.sharePoint.properties.pane.content[this.key].settings.myInstagram.instaCaption = true;
					this.sharePoint.properties.pane.content[this.key].draft.newEndPoint = finalURL + `&amp;omitscript=true&amp;hidecaption=true${finalWidth}`;
					break;
				case 'show':
					this.sharePoint.properties.pane.content[this.key].settings.myInstagram.instaCaption = false;
					this.sharePoint.properties.pane.content[this.key].draft.newEndPoint = finalURL + `&amp;omitscript=true${finalWidth}`;
					break;
			}
		});

		let changeWidth = instagramPane.find('#width-cell');
		changeWidth.onChanged(value => {
			this.sharePoint.properties.pane.content[this.key].settings.myInstagram.instaWidth = `&amp;minwidth=${value}&amp;maxwidth=${value}`;
			let finalURL = this.sharePoint.properties.pane.content[this.key].settings.myInstagram.instaURL || this.defaultURL;
			const finalHide = (this.sharePoint.properties.pane.content[this.key].settings.myInstagram.instaCaption) ? '&amp;hidecaption=true' : '';

			this.sharePoint.properties.pane.content[this.key].draft.newEndPoint = finalURL + `&amp;omitscript=true${finalHide}${this.sharePoint.properties.pane.content[this.key].settings.myInstagram.instaWidth}`;

		});


		this.paneContent.addEventListener('mutated', event => {
			// this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
			this.getEmbed(this.sharePoint.properties.pane.content[this.key].draft.newEndPoint);
			this.element.innerHTML = draftDom.innerHTML;//upate the webpart
			this.element.css(draftDom.css());
		});
	}
}

class YouTube extends CraterWebParts {
	public params;
	public element;
	public key;
	public paneContent;
	public elementModifier = new ElementModifier();
	public player;

	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.params = params;
	}

	public render(params) {

		let youtube = this.createKeyedElement({
			element: 'div', attributes: { class: 'crater-youtube crater-component', 'data-type': 'youtube' }, children: [
				{
					element: 'div', attributes: { class: 'crater-youtube-contents' }, children: [
						{
							element: 'div', attributes: { class: 'crater-iframe' }, children: [
								{ element: 'iframe', attributes: { id: 'player', type: 'text/html', width: '640', height: '390', src: `https://www.youtube.com/embed/M7lc1UVf-VE?origin=${location.href}&autoplay=1&enablejsapi=1&widgetid=1`, frameborder: '0' } }
							]
						}
					]
				}
			]
		});

		this.key = youtube.dataset.key;
		this.sharePoint.properties.pane.content[this.key].settings = { myYoutube: { defaultVideo: 'https://www.youtube.com/embed/M7lc1UVf-VE', width: '640', height: '390' } };

		let youtubeContent = youtube.find('.crater-youtube-contents');
		youtubeContent.makeElement({
			element: 'script', attributes: { src: 'https://www.youtube.com/iframe_api' }
		});

		return youtube;
	}

	public rendered(params) { }

	public setUpPaneContent(params) {
		let key = params.element.dataset['key'];
		this.element = this.sharePoint.properties.pane.content[key].draft.dom;
		this.paneContent = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'crater-property-content' }
		}).monitor();

		if (this.sharePoint.properties.pane.content[key].draft.pane.content.length !== 0) {//check if draft pane content is not empty and set it to the pane content
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[key].draft.pane.content;
		}
		else if (this.sharePoint.properties.pane.content[key].content.length !== 0) {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[key].content;
		} else {
			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'youtube-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Customise Youtube'
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							{
								element: 'div', attributes: { class: 'message-note' }, children: [
									{ element: 'span', attributes: { id: 'videoURL', style: { color: 'green' } }, text: `Note: The Video URL is in this format "https://youtu.be/SHoBUYvsjsc"` },
								]
							}]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'videoURL', value: this.sharePoint.properties.pane.content[key].settings.myYoutube.defaultVideo
							}),
							this.elementModifier.cell({
								element: 'input', name: 'width', value: this.sharePoint.properties.pane.content[key].settings.myYoutube.width
							}),
							this.elementModifier.cell({
								element: 'input', name: 'height', value: this.sharePoint.properties.pane.content[key].settings.myYoutube.height
							})
						]
					})
				]
			});
		}
		return this.paneContent;
	}

	public generatePaneContent(params) {

	}

	public listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;

		this.paneContent = this.sharePoint.app.find('.crater-property-content').monitor();

		let youtubePane = this.paneContent.find('.youtube-pane');

		youtubePane.find('#videoURL-cell').onChanged(value => {
			if (value.indexOf('.be/') !== -1) {
				youtubePane.find('#videoURL').style.color = 'green';
				youtubePane.find('#videoURL').textContent = 'Valid URL';
				let afterEmbed = value.split('.be/')[1];
				let newValue = 'https://www.youtube.com/embed/' + afterEmbed;
				this.sharePoint.properties.pane.content[this.key].settings.myYoutube.defaultVideo = newValue;
			} else {
				youtubePane.find('#videoURL').style.color = 'red';
				youtubePane.find('#videoURL').textContent = 'Invalid Video URL. Please right click on the video to get the video URL';
			}

		});

		youtubePane.find('#width-cell').onChanged(value => {
			this.sharePoint.properties.pane.content[this.key].settings.myYoutube.width = value;
		});

		youtubePane.find('#height-cell').onChanged(value => {
			this.sharePoint.properties.pane.content[this.key].settings.myYoutube.height = value;
		});

		this.paneContent.addEventListener('mutated', event => {
			// this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {

			let draftDomIframe = draftDom.find('.crater-iframe');

			draftDomIframe.innerHTML = '';

			draftDomIframe.makeElement({
				element: 'iframe', attributes: { id: 'player2', type: 'text/html', width: this.sharePoint.properties.pane.content[this.key].settings.myYoutube.width, height: this.sharePoint.properties.pane.content[this.key].settings.myYoutube.height, src: `${this.sharePoint.properties.pane.content[this.key].settings.myYoutube.defaultVideo}?origin=${location.href}&autoplay=1&enablejsapi=1&widgetid=1`, frameborder: '0' }
			});


			console.log(draftDom);
			this.element.css(draftDom.css());
			// this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;
			this.element.innerHTML = draftDom.innerHTML;//upate the webpart
		});

	}
}

class Facebook extends CraterWebParts {
	private params: any;
	public key: any;
	public elementModifier = new ElementModifier();
	public pageURL: any = `https://www.facebook.com/ipisolutionsnigerialtd`;
	public element: any;
	public paneContent: any;

	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.params = params;
	}

	public render(params) {
		let facebookDiv = this.createKeyedElement({
			element: 'div', attributes: { class: 'crater-facebook crater-component', 'data-type': 'facebook' }
		});

		this.key = facebookDiv.dataset.key;
		this.element = facebookDiv;
		this.sharePoint.properties.pane.content[this.key].settings.facebook = { url: this.pageURL, tabs: 'timeline,messages,events', smallHeader: 'false', hideCover: 'false', showFacePile: 'false' };
		return facebookDiv;
	}

	public rendered(params) {
		this.element = params.element;
		this.key = params.element.dataset['key'];
		let facebook = this.sharePoint.properties.pane.content[this.key].settings.facebook;
		this.faceBookSettings({ url: facebook.url, dataTabs: facebook.tabs, smallHeader: facebook.smallHeader, hideCover: facebook.hideCover, showFacePile: facebook.showFacePile });
	}

	public faceBookSettings(params) {
		window.onerror = (msg, url, lineNumber, columnNumber, error) => {
			console.log(msg, url, lineNumber, columnNumber, error);
		};
		try {
			this.key = this.element.dataset['key'];
			const width = (params.width) ? params.width : this.element.getBoundingClientRect().width;
			const height = (params.height) ? params.height : '';
			const adaptContainer = (params.adaptContainer) ? params.adaptContainer : 'true';
			let crater = this.element.find('.crater-facebook-content');
			if (crater) crater.remove();

			let facebookContent = this.elementModifier.createElement({
				element: 'div', attributes: { class: 'crater-facebook-content' }
			});
			this.element.appendChild(facebookContent);

			facebookContent.makeElement({
				element: 'div', attributes: { id: 'fb-root' }
			});
			facebookContent.makeElement({
				element: 'script', attributes: { class: "facebook-script", src: 'https://connect.facebook.net/en_US/sdk.js#xfbml=1&version=v5.0&appId=541045216450969&autoLogAppEvents=1', async: true, defer: true }
			});
			facebookContent.makeElement({
				element: 'div', attributes: {
					class: 'fb-page',
					'data-href': params.url, 'data-tabs': params.dataTabs, 'data-width': width, 'data-height': height, 'data-small-header': params.smallHeader, 'data-adapt-container-width': adaptContainer, 'data-hide-cover': params.hideCover, 'data-show-facepile': params.showFacePile
				}
			});
			//@ts-ignore
			FB.init({
				appId: '541045216450969',
				autoLogAppEvents: true,
				xfbml: true,
				version: 'v5.0'
			});

		} catch (error) {
			console.log(error.message);
		}

	}

	public setUpPaneContent(params) {
		this.key = params.element.dataset['key'];
		this.element = this.sharePoint.properties.pane.content[this.key].draft.dom;
		this.paneContent = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'crater-property-content' }
		});

		if (this.sharePoint.properties.pane.content[this.key].draft.pane.content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.pane.content;

		}
		else if (this.sharePoint.properties.pane.content[this.key].content != '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[this.key].content;
		} else {
			// console.log(params)
			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'title-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' },
						children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: "title" }, text: 'Page'
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' },
						children: [
							this.elementModifier.cell({
								element: 'input', name: 'pageUrl',
								value:
									//this.defaultURL
									this.sharePoint.properties.pane.content[this.key].draft.dom.find('.fb-page ').dataset.href
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Tabs', value: this.sharePoint.properties.pane.content[this.key].settings.facebook.tabs
							}),
							this.elementModifier.cell({
								element: 'select', name: 'hide-cover', options: ['show', 'hide']
							}),
							this.elementModifier.cell({
								element: 'select', name: 'hide-facepile', options: ['show', 'hide']
							}),
							this.elementModifier.cell({
								element: 'select', name: 'small-header', options: ['show', 'hide']
							}),
						]
					})
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'size-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Size'
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							{
								element: 'div', attributes: { class: 'message-note' }, children: [
									{
										element: 'div', attributes: { class: 'message-text' }, children: [
											{ element: 'p', attributes: { style: { color: 'green' } }, text: `Adapt to container width: Fit the component to parent container's width. The minimum and maximum width values below still apply` },
											{ element: 'p', attributes: { style: { color: 'green' } }, text: `Width: min:180 max: 500 pixels` },
											{ element: 'p', attributes: { style: { color: 'green' } }, text: `Height: min: 70 max: 1600 pixels` }
										]
									}
								]
							}]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'select', name: 'container-width', options: ['true', 'false']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'width', value: this.sharePoint.properties.pane.content[this.key].settings.facebook.width || ''
							}),
							this.elementModifier.cell({
								element: 'input', name: 'height', value: this.sharePoint.properties.pane.content[this.key].settings.facebook.height || ''
							})
						]
					})
				]
			});
		}

		return this.paneContent;
	}

	public generatePaneContent(params) { }

	public listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		this.paneContent = this.sharePoint.app.find(".crater-property-content").monitor();

		let titlePane = this.paneContent.find('.title-pane');
		let sizePane = this.paneContent.find('.size-pane');
		let facebook = this.sharePoint.properties.pane.content[this.key].settings.facebook;


		titlePane.find('#pageUrl-cell').onChanged(value => {
			facebook.url = value;
		});
		titlePane.find('#Tabs-cell').onChanged(value => {
			facebook.tabs = value;
		});
		let coverCell = titlePane.find('#hide-cover-cell');
		coverCell.addEventListener('change', e => {
			switch (coverCell.value.toLowerCase()) {
				case 'show':
					facebook.hideCover = 'true';
					break;
				case 'hide':
					facebook.hideCover = 'false';
					break;
				default:
					facebook.hideCover = 'false';
			}
		});

		let hideFacePileCell = titlePane.find('#hide-facepile-cell');
		hideFacePileCell.addEventListener('change', e => {
			switch (hideFacePileCell.value.toLowerCase()) {
				case 'show':
					facebook.showFacePile = 'true';
					break;
				case 'hide':
					facebook.showFacePile = 'false';
					break;
				default:
					facebook.showFacePile = 'false';
			}
		});

		let showHeaderCell = titlePane.find('#small-header-cell');
		showHeaderCell.addEventListener('change', e => {
			switch (showHeaderCell.value.toLowerCase()) {
				case 'show':
					facebook.smallHeader = 'true';
					break;
				case 'hide':
					facebook.smallHeader = 'false';
					break;
				default:
					facebook.smallHeader = 'false';
			}
		});

		sizePane.find('#width-cell').onChanged(value => {
			facebook.width = value;
		});

		sizePane.find('#height-cell').onChanged(value => {
			facebook.height = value;
		});

		let adaptCell = sizePane.find('#container-width-cell');
		adaptCell.addEventListener('change', e => {
			switch (adaptCell.value.toLowerCase()) {
				case 'true':
					facebook.adaptContainer = 'true';
					break;
				case 'false':
					facebook.adaptContainer = 'false';
					break;
				default:
					facebook.adaptContainer = 'false';
			}
		});
		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {

			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;

			this.faceBookSettings({ url: facebook.url, dataTabs: facebook.tabs, smallHeader: facebook.smallHeader, hideCover: facebook.hideCover, showFacePile: facebook.showFacePile, width: facebook.width, height: facebook.height, adaptContainer: facebook.adaptContainer });
		});
	}
}

class BeforeAfter extends CraterWebParts {
	public params: any;
	public elementModifier = new ElementModifier();
	public paneContent: any;
	public element: any;
	private key: any;

	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.params = params;

	}

	public render(params) {
		let beforeAfterDiv = this.createKeyedElement({
			element: 'div', attributes: { class: 'crater-beforeAfter crater-component', 'data-type': 'beforeafter' }, children: [
				this.elementModifier.createElement({
					element: 'div', attributes: { class: 'crater-beforeAfter-contents' }, children: [
						{
							element: 'img', attributes: { class: 'crater-beforeImage', src: "http://egegorgulu.com/assets/img/beforeafter/before.jpg" }
						},

						{
							element: 'div', attributes: { class: 'crater-after' }, children: [
								{
									element: 'img', attributes: { class: 'crater-afterImage', src: "http://egegorgulu.com/assets/img/beforeafter/after.jpg" }
								}
							]
						},
						{
							element: 'span', attributes: { class: 'crater-handle' }
						}
					]
				})
			]
		});

		return beforeAfterDiv;
	}

	public rendered(params) {
		this.element = params.element;
		const slider = this.element.find('.crater-beforeAfter-contents').find('.crater-handle');
		let isDown = false;
		let resizeDiv = this.element.find('.crater-beforeAfter-contents').find('.crater-after');
		let containerWidth = this.element.find('.crater-beforeAfter-contents').offsetWidth + 'px';
		this.element.find('.crater-after img').css({ "width": containerWidth });

		this.drags(slider, resizeDiv, this.element.find('.crater-beforeAfter-contents'));
	}

	private drags(dragElement, resizeElement, container): any {
		//initialize the dragging event on mousedown
		dragElement.addEventListener('mousedown', elementDown => {

			dragElement.classList.add('crater-draggable');
			resizeElement.classList.add('crater-resizable');
			let startX = elementDown.pageX;
			//get the initial position
			let dragWidth = dragElement.clientWidth,
				posX = dragElement.offsetLeft + dragWidth - startX,
				containerOffset = container.offsetLeft,
				containerWidth = container.clientWidth;
			// Set limits
			let minLeft = containerOffset + 10;
			let maxLeft = containerOffset + containerWidth - dragWidth - 10;

			dragElement.parentNode.addEventListener('mousemove', elementMoved => {
				let moveX = elementMoved.pageX;
				let leftValue = moveX + posX - dragWidth;

				// Prevent going off limits
				if (leftValue < minLeft) {
					leftValue = minLeft;
				} else if (leftValue > maxLeft) {
					leftValue = maxLeft;
				}
				// Translate the handle's left value to masked divs width.
				var widthValue = (leftValue + dragWidth / 2 - containerOffset) * 100 / containerWidth + '%';
				// Set the new values for the slider and the handle. 
				// Bind mouseup events to stop dragging.

				let draggable = container.find('.crater-draggable');

				if (!func.isnull(draggable)) draggable.css({ 'left': widthValue });
				container.addEventListener('mouseup', function () {
					this.classList.remove('crater-draggable');
					resizeElement.classList.remove('crater-resizable');
				});

				let resizable = container.find('.crater-resizable');
				if (!func.isnull(resizable)) resizable.css({ 'width': widthValue });
			});
			container.addEventListener('mouseup', () => {
				dragElement.classList.remove('crater-draggable');
				resizeElement.classList.remove('crater-resizable');
			});
			elementDown.preventDefault();
		});
		dragElement.addEventListener('mouseup', elementUp => {
			dragElement.classList.remove('crater-draggable');
			resizeElement.classList.remove('crater-resizable');
		});
	}

	public setUpPaneContent(params) {
		let key = params.element.dataset['key'];
		this.element = this.sharePoint.properties.pane.content[key].draft.dom;
		this.paneContent = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'crater-property-content' }
		}).monitor();
		if (this.sharePoint.properties.pane.content[key].draft.pane.content! = "") {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[key].draft.pane.content;
		}
		else if (this.sharePoint.properties.pane.content[key].content != "") {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[key].content;
		}
		else {

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'title-pane card' }, children: [
					{
						element: 'div', attributes: { class: 'card-title' }, children: [
							{ element: 'h2', attributes: { class: 'title' }, text: 'Settings' }
						]
					},
					{
						element: 'div', attributes: { class: 'row' },
						children: [
							this.elementModifier.cell({
								element: "img", name: "before",
								dataAttributes: {
									style: { width: '400px', height: '400px' },
									src: this.element.find('.crater-beforeAfter-contents').find('.crater-beforeImage').src
								}
							}),
							this.elementModifier.cell({
								element: 'img',
								name: 'after',
								dataAttributes: {
									style: { width: '400px', height: '400px' },
									src: this.element.find('.crater-beforeAfter-contents').find('.crater-after').find('.crater-afterImage').src
								}
							})
						]
					}
				]
			});
		}
		return this.paneContent;
	}

	public listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		this.paneContent = this.sharePoint.app.find('.crater-property-content').monitor();

		let afterCell = this.paneContent.find('#after-cell').parentNode;
		this.uploadImage({ parent: afterCell }, (image) => {
			afterCell.find('#after-cell').src = image.src;
			this.sharePoint.properties.pane.content[this.key].draft.dom.find('.beforeAfter-contents').find('.crater-after').find('.afterImage').src = image.src;
		});

		let beforeCell = this.paneContent.find('#before-cell').parentNode;
		this.uploadImage({ parent: beforeCell }, (image) => {
			beforeCell.find('#before-cell').src = image.src;
			this.sharePoint.properties.pane.content[this.key].draft.dom.find('.beforeAfter-contents').find('.beforeImage').src = image.src;
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());

			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart
		});
	}
}

class Event extends CraterWebParts {
	public params;
	public element;
	public key;
	public paneContent;
	public elementModifier = new ElementModifier();
	public today: any;
	public month: any;
	public day: any;
	public monthArray = func.trimMonthArray();


	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.params = params;
	}


	public render(params) {
		this.today = func.today();
		this.month = func.isMonthValid(this.today);
		this.day = func.isDayValid(this.today);

		if (!func.isset(params.source)) {
			params.source = [
				{
					icon: 'https://img.icons8.com/pastel-glyph/64/000000/christmas-tree.png',
					title: "BasketBall Game",
					location: "Lagos Island, Lagos",
					day: "19",
					month: "Aug",
					start: '12:00AM',
					end: '01:00AM'
				},
				{
					icon: 'https://img.icons8.com/cute-clipart/64/000000/shoes.png',
					title: "Shoe City Event",
					location: "Ikeja, Lagos",
					day: "28",
					month: "oct",
					start: '01:00PM',
					end: '02:00PM'
				},
				{
					icon: 'https://img.icons8.com/cotton/64/000000/football-ball.png',
					title: "Football Game",
					location: "Teslim Balogun Statdium, Lagos",
					day: "21",
					month: "Jan",
					start: '02:00PM',
					end: '03:00PM'
				}
			];
		}

		let event = this.createKeyedElement({
			element: 'div', attributes: { class: 'crater-event', 'data-type': 'event' }, children: [
				{
					element: 'div', attributes: { class: 'crater-event-title' }, children: [
						{ element: 'img', attributes: { class: 'crater-event-title-imgIcon', src: 'https://img.icons8.com/cute-clipart/64/000000/tear-off-calendar.png' } },
						{ element: 'span', attributes: { class: 'crater-event-title-captionTitle' }, text: 'Events' }
					]
				},
				{ element: 'div', attributes: { class: 'crater-event-content' } }
			]
		});

		let content = event.find(`.crater-event-content`);
		let locationElement = content.findAll('.crater-event-content-task-location') as any;

		for (let each of params.source) {
			content.makeElement(
				{
					element: 'div', attributes: { class: 'crater-event-content-item' }, children: [
						this.elementModifier.createElement({
							element: 'div', attributes: { class: 'crater-event-content-item-icon' }, children: [
								this.elementModifier.createElement({ element: 'img', attributes: { class: 'crater-event-content-item-icon-image', id: 'icon', src: func.isnull(each.icon) ? this.sharePoint.images.append : each.icon, alt: 'Event Icon' } })
							]
						}),
						this.elementModifier.createElement({
							element: 'div', attributes: { class: 'crater-event-content-task' }, children: [
								{ element: 'div', attributes: { class: 'crater-event-content-task-caption', id: 'eventTask' }, text: each.title },
								{
									element: 'div', attributes: { class: 'crater-event-content-task-location' }, children: [
										{ element: 'img', attributes: { src: 'https://img.icons8.com/small/16/000000/clock.png' } },
										{ element: 'span', attributes: { class: 'crater-event-content-task-location-duration', id: 'startTime' }, text: `${each.start} - ` },
										{ element: 'span', attributes: { class: 'crater-event-content-task-location-duration', id: 'endTime' }, text: `${each.end}` },
										{ element: 'img', attributes: { src: 'https://img.icons8.com/small/16/000000/previous--location.png' } },
										{ element: 'span', attributes: { class: 'crater-event-content-task-location-place', id: 'location' }, text: `${each.location}` }
									]
								}
							]
						}),
						{
							element: 'div', attributes: { class: 'crater-event-content-item-date' }, children: [
								{
									element: 'div', attributes: { class: 'crater-event-content-item-date-day', id: 'Day' }, text: each.day
								},
								{ element: 'div', attributes: { class: 'crater-event-content-item-date-month', id: 'Month' }, text: each.month.toUpperCase() }
							]
						}
					]
				}
			);
		}
		this.key = this.key || event.dataset.key;
		return event;
	}

	public rendered(params) {
		this.element = params.element;

	}

	public setUpPaneContent(params) {
		let key = params.element.dataset['key'];//create a key variable and set it to the webpart key
		this.element = this.sharePoint.properties.pane.content[key].draft.dom;//define the declared element to the draft dom content
		this.paneContent = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'crater-property-content' }
		}).monitor(); //monitor the content pane 
		if (this.sharePoint.properties.pane.content[key].draft.pane.content !== '') {//check if draft pane content is not empty and set it to the pane content
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[key].draft.pane.content;
		}
		else if (this.sharePoint.properties.pane.content[key].content !== '') {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[key].content;
		} else {
			let eventList = this.sharePoint.properties.pane.content[key].draft.dom.find('.crater-event-content');
			let dateListRows = eventList.findAll('.crater-event-content-item');
			this.paneContent.makeElement({
				element: 'div', children: [
					this.elementModifier.createElement(
						{ element: 'button', attributes: { class: 'btn new-component', style: { display: 'inline-block', borderRadius: '5px' } }, text: 'Add New' }
					)
				]
			});


			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'title-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							{ element: 'h2', attributes: { class: 'title' }, text: 'Event Title Layout' }
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [//create the cells for changing crater event title
							this.elementModifier.cell({
								element: 'img', name: 'icon', attributes: {}, dataAttributes: { class: 'crater-icon', src: this.element.find('.crater-event-title-imgIcon').src }
							}),
							this.elementModifier.cell({
								element: 'input', name: 'title', value: this.element.find('.crater-event-title').textContent
							}),
							this.elementModifier.cell({
								element: 'input', name: 'backgroundcolor', value: this.element.find('.crater-event-title').css()['background-color']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'color', value: this.element.find('.crater-event-title').css().color
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontsize', value: this.element.find('.crater-event-title').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'height', value: this.element.find('.crater-event-title').css()['height'] || ''
							}),
							this.elementModifier.cell({
								element: 'select', name: 'toggleTitle', options: ['show', 'hide']
							})
						]
					})
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'event-icon-row-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Edit Event Icon'
							})
						]
					}),
					{
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'iconWidth', value: this.element.find('.crater-event-content-item-icon-image').css()['width']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'iconHeight', value: this.element.find('.crater-event-content-item-icon-image').css()['height']
							}),
							this.elementModifier.cell({
								element: 'select', name: 'toggleIcon', options: ['show', 'hide']
							})
						]
					}
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'event-title-row-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Edit Event Title'
							})
						]
					}),
					{
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'fontSize', value: this.element.find('.crater-event-content-task-caption').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontFamily', value: this.element.find('.crater-event-content-task-caption').css()['font-family']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'eventColor', value: this.element.find('.crater-event-content-task-caption').css()['color']
							}),
							this.elementModifier.cell({
								element: 'select', name: 'toggleTitle', options: ['show', 'hide']
							})
						]
					}
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'event-location-row-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Edit Event Location'
							})
						]
					}),
					{
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'fontSize', value: this.element.find('.crater-event-content-task-location-place').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontFamily', value: this.element.find('.crater-event-content-task-location-place').css()['font-family']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'locationColor', value: this.element.find('.crater-event-content-task-location-place').css().color
							}),
							this.elementModifier.cell({
								element: 'select', name: 'toggleLocation', options: ['show', 'hide']
							})
						]
					}
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'event-duration-row-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Edit Event Duration'
							})
						]
					}),
					{
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'fontSize', value: this.element.find('.crater-event-content-task-location-duration').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontFamily', value: this.element.find('.crater-event-content-task-location-duration').css()['font-family']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'durationColor', value: this.element.find('.crater-event-content-task-location-duration').css()['color']
							}),
							this.elementModifier.cell({
								element: 'select', name: 'toggleDuration', options: ['show', 'hide']
							})
						]
					}
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'event-date-row-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Edit Event Date '
							})
						]
					}),
					{
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'daySize', value: this.element.find('.crater-event-content-item-date-day').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'monthSize', value: this.element.find('.crater-event-content-item-date-month').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontFamily', value: this.element.find('.crater-event-content-item-date').css()['font-family']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'dateColor', value: this.element.find('.crater-event-content-item-date').css()['color']
							}),
							this.elementModifier.cell({
								element: 'select', name: 'toggleDate', options: ['show', 'hide']
							})
						]
					}
				]
			});

			this.paneContent.append(this.generatePaneContent({ list: dateListRows }));

		}
		return this.paneContent;
	}


	public generatePaneContent(params) {
		let eventListPane = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'card list-pane' }, children: [
				this.elementModifier.createElement({
					element: 'div', attributes: { class: 'card-title' }, children: [
						this.elementModifier.createElement({
							element: 'h2', attributes: { class: 'title' }, text: 'Events List'
						})
					]
				}),
			]
		});

		let strip = (value) => {
			return value.split(' ');
		};
		let cTime = strip(this.element.find('#startTime').textContent)[0];
		let dTime = this.element.find('#endTime').textContent;
		// let cDay = this.element.find('.crater-event-content-item-date-day').textContent;
		// let cMonth = this.element.find('.crater-event-content-item-date-month').textContent;
		// let gDate = new Date(`${cMonth} ${cDay}, 2019`);


		for (let i = 0; i < params.list.length; i++) {
			eventListPane.makeElement({
				element: 'div',
				attributes: { class: 'crater-event-item-pane row' },
				children: [
					this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-event-content-item' }),
					this.elementModifier.cell({
						element: 'img', name: 'icon', attributes: {}, dataAttributes: { class: 'crater-icon', src: params.list[i].find('#icon').src }
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Task', attributes: { class: 'taskValue' }, value: params.list[i].find('#eventTask').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Location', attributes: { class: 'locationValue' }, value: params.list[i].find('#location').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Day', attribute: { class: 'crater-date dateValue' }, value: params.list[i].find('#Day').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Month', attribute: { class: 'crater-date dateValue' }, value: params.list[i].find('#Month').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'start', attributes: { class: 'startValue' }, value: cTime
					}),
					this.elementModifier.cell({
						element: 'input', name: 'end', attributes: { class: 'endValue' }, value: dTime
					})
				]
			});
		}

		return eventListPane;

	}

	public listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		this.paneContent = this.sharePoint.app.find('.crater-property-content').monitor();
		//get the content and all the events
		let eventList = this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-event-content');
		let eventListRow = eventList.findAll('.crater-event-content-item');

		let eventListRowPanePrototype = this.elementModifier.createElement({//create a row on the property pane
			element: 'div',
			attributes: {
				class: 'crater-event-item-pane row'
			},
			children: [
				this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-event-content-item' }),
				this.elementModifier.cell({
					element: 'input', name: 'icon', value: this.sharePoint.images.append
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Task', value: 'New Task'
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Location', value: 'Lagos, Nigeria'
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Day', attributes: { placeholder: 'e.g 1' },
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Month', attributes: { placeholder: 'preferably first three letters only e.g Jan' },
				}),
				this.elementModifier.cell({
					element: 'input', name: 'start', dataAttributes: { type: 'time' }
				}),
				this.elementModifier.cell({
					element: 'input', name: 'end', dataAttributes: { type: 'time' }
				})
			]
		});


		let eventListRowDomPrototype = this.createKeyedElement(
			{
				element: 'div', attributes: { class: 'crater-event-content-item keyed-element' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'crater-event-content-item-icon' }, children: [
							this.elementModifier.createElement({ element: 'img', attributes: { class: 'crater-event-content-item-icon-image', src: this.sharePoint.images.append, alt: 'Event Icon' } })
						]
					}),
					{
						element: 'div', attributes: { class: 'crater-event-content-task' }, children: [
							{ element: 'div', attributes: { class: 'crater-event-content-task-caption', id: 'eventTask' }, text: this.paneContent.find(`.taskValue input`).value },
							{
								element: 'div', attributes: { class: 'crater-event-content-task-location' }, children: [
									{ element: 'img', attributes: { src: 'https://img.icons8.com/small/16/000000/clock.png' } },
									{ element: 'span', attributes: { class: 'crater-event-content-task-location-duration startTime' }, text: '' },
									{ element: 'span', attributes: { class: 'crater-event-content-task-location-duration endTime' }, text: '' },
									{ element: 'img', attributes: { src: 'https://img.icons8.com/small/16/000000/previous--location.png' } },
									{ element: 'span', attributes: { class: 'crater-event-content-task-location-place' }, text: this.paneContent.find(`.locationValue input`).value }
								]
							}
						]
					},
					{
						element: 'div', attributes: { class: 'crater-event-content-item-date' }, children: [
							{ element: 'div', attributes: { class: 'crater-event-content-item-date-day', id: 'Day' }, text: '' },
							{ element: 'div', attributes: { class: 'crater-event-content-item-date-month', id: 'Month' }, text: '' }
						]
					},
				]
			}
		);


		let eventRowHandler = (eventRowPane, eventRowDom) => {
			eventRowPane.addEventListener('mouseover', event => {
				eventRowPane.find('.crater-content-options').css({ visibility: 'visible' });
			});

			eventRowPane.addEventListener('mouseout', event => {
				eventRowPane.find('.crater-content-options').css({ visibility: 'hidden' });
			});

			let iconCellParent = eventRowPane.find('#icon-cell').parentNode;
			this.uploadImage({ parent: iconCellParent }, (image) => {
				iconCellParent.find('#icon-cell').src = image.src;
				eventRowDom.find('.crater-event-content-item-icon-image').src = image.src;
			});

			// get the values of the newly created row on the property - pane
			eventRowPane.find('#Task-cell').onChanged(value => {
				eventRowDom.find('.crater-event-content-task-caption').innerHTML = value;
			});

			eventRowPane.find('#Location-cell').onChanged(value => {
				eventRowDom.find('.stateCountry').innerHTML = value;
			});

			eventRowPane.find('#Day-cell').onChanged(value => {
				eventRowDom.find('.crater-event-content-item-date-day').innerHTML = value;
			});

			eventRowPane.find('#Month-cell').onChanged(value => {
				eventRowDom.find('.crater-event-content-item-date-month').innerHTML = value;
			});

			eventRowPane.find('#start-cell').onChanged(value => {
				eventRowDom.find('.startTime').innerHTML = value + ` - `;
			});

			eventRowPane.find('#end-cell').onChanged(value => {
				eventRowDom.find('.endTime').innerHTML = value;
			});

			eventRowPane.find('.delete-crater-event-content-item').addEventListener('click', event => {
				eventRowDom.remove();
				eventRowPane.remove();
			});

			eventRowPane.find('.add-before-crater-event-content-item').addEventListener('click', event => {
				let newEventRowDom = eventListRowDomPrototype.cloneNode(true);
				let neweventRowPane = eventListRowPanePrototype.cloneNode(true);

				eventRowDom.before(newEventRowDom);
				eventRowPane.before(neweventRowPane);
				eventRowHandler(neweventRowPane, newEventRowDom);
			});

			eventRowPane.find('.add-after-crater-event-content-item').addEventListener('click', event => {
				let newEventRowDom = eventListRowDomPrototype.cloneNode(true);
				let newEventRowPane = eventListRowPanePrototype.cloneNode(true);

				eventRowDom.after(newEventRowDom);
				eventRowPane.after(newEventRowPane);

				eventRowHandler(newEventRowPane, newEventRowDom);
			});
		};

		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;

		let titlePane = this.paneContent.find('.title-pane');
		let eventIconRowPane = this.paneContent.find('.event-icon-row-pane');
		let eventTitleRowPane = this.paneContent.find('.event-title-row-pane');
		let eventLocationRowPane = this.paneContent.find('.event-location-row-pane');
		let eventDurationRowPane = this.paneContent.find('.event-duration-row-pane');
		let eventDateRowPane = this.paneContent.find('.event-date-row-pane');

		let iconParent = titlePane.find('#icon-cell').parentNode;

		let eventColorParent = eventTitleRowPane.find('#eventColor-cell').parentNode;
		this.pickColor({ parent: eventColorParent, cell: eventColorParent.find('#eventColor-cell') }, (color) => {
			draftDom.findAll('.crater-event-content-task-caption').forEach(element => {
				element.css({ color });//get the color of the event font in the draftDom
			});
			eventColorParent.find('#eventColor-cell').value = color;
			eventColorParent.find('#eventColor-cell').setAttribute('value', color); //set the value of the eventColor cell in the pane to the color
		});

		let dateColorParent = eventDateRowPane.find('#dateColor-cell').parentNode;
		this.pickColor({ parent: dateColorParent, cell: dateColorParent.find('#dateColor-cell') }, (color) => {
			draftDom.findAll('.crater-event-content-item-date').forEach(element => {
				element.css({ color });//get the color of the event font in the draftDom
			});
			dateColorParent.find('#dateColor-cell').value = color;
			dateColorParent.find('#dateColor-cell').setAttribute('value', color); //set the value of the eventColor cell in the pane to the color
		});

		let locationColorParent = eventLocationRowPane.find('#locationColor-cell').parentNode;
		this.pickColor({ parent: locationColorParent, cell: locationColorParent.find('#locationColor-cell') }, (color) => {
			draftDom.findAll('.crater-event-content-task-location-place').forEach(element => {
				element.css({ color });//get the color of the event font in the draftDom
			});
			locationColorParent.find('#locationColor-cell').value = color;
			locationColorParent.find('#locationColor-cell').setAttribute('value', color); //set the value of the eventColor cell in the pane to the color
		});

		let durationColorParent = eventDurationRowPane.find('#durationColor-cell').parentNode;
		this.pickColor({ parent: durationColorParent, cell: durationColorParent.find('#durationColor-cell') }, (color) => {
			draftDom.findAll('.crater-event-content-task-location-duration').forEach(element => {
				element.css({ color });//get the color of the event font in the draftDom
			});
			durationColorParent.find('#durationColor-cell').value = color;
			durationColorParent.find('#durationColor-cell').setAttribute('value', color); //set the value of the eventColor cell in the pane to the color
		});
		this.uploadImage({ parent: iconParent }, (image) => {
			iconParent.find('#icon-cell').src = image.src;
			draftDom.find('.crater-event-title-imgIcon').src = image.src;
		});
		titlePane.find('#title-cell').onChanged(value => {
			draftDom.find('.crater-event-title-captionTitle').innerHTML = value;
		});

		let backgroundColorCell = titlePane.find('#backgroundcolor-cell').parentNode;
		this.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.find('#backgroundcolor-cell') }, (backgroundColor) => {
			draftDom.find('.crater-event-title').css({ backgroundColor });
			backgroundColorCell.find('#backgroundcolor-cell').value = backgroundColor;
			backgroundColorCell.find('#backgroundcolor-cell').setAttribute('value', backgroundColor);
		});

		let colorCell = titlePane.find('#color-cell').parentNode;
		this.pickColor({ parent: colorCell, cell: colorCell.find('#color-cell') }, (color) => {
			draftDom.find('.crater-event-title').css({ color });
			colorCell.find('#color-cell').value = color;
			colorCell.find('#color-cell').setAttribute('value', color);
		});


		titlePane.find('#fontsize-cell').onChanged(value => {
			draftDom.find('.crater-event-title').css({ fontSize: value });
		});

		titlePane.find('#height-cell').onChanged(value => {
			draftDom.find('.crater-event-title').css({ height: value });
		});

		eventIconRowPane.find('#iconWidth-cell').onChanged(value => {
			draftDom.findAll('.crater-event-content-item-icon-image').forEach(element => {
				element.css({ width: value });
			});
		});
		eventIconRowPane.find('#iconHeight-cell').onChanged(value => {
			draftDom.findAll('.crater-event-content-item-icon-image').forEach(element => {
				element.css({ height: value });
			});
		});

		eventTitleRowPane.find('#fontSize-cell').onChanged(value => {
			draftDom.findAll('.crater-event-content-task-caption').forEach(element => {
				element.css({ fontSize: value });
			});
		});
		eventTitleRowPane.find('#fontFamily-cell').onChanged(value => {
			draftDom.findAll('.crater-event-content-task-caption').forEach(element => {
				element.css({ fontFamily: value });
			});
		});

		eventLocationRowPane.find('#fontSize-cell').onChanged(value => {
			draftDom.findAll('.crater-event-content-task-location-place').forEach(element => {
				element.css({ fontSize: value });
			});
		});
		eventLocationRowPane.find('#fontFamily-cell').onChanged(value => {
			draftDom.findAll('.crater-event-content-task-location-place').forEach(element => {
				element.css({ fontFamily: value });
			});
		});

		eventDurationRowPane.find('#fontSize-cell').onChanged(value => {
			draftDom.findAll('.crater-event-content-task-location-duration').forEach(element => {
				element.css({ fontSize: value });
			});
		});
		eventDurationRowPane.find('#fontFamily-cell').onChanged(value => {
			draftDom.findAll('.crater-event-content-task-location-duration').forEach(element => {
				element.css({ fontFamily: value });
			});
		});

		eventDateRowPane.find('#daySize-cell').onChanged(value => {
			draftDom.findAll('.crater-event-content-item-date-day').forEach(element => {
				element.css({ fontSize: value });
			});
		});
		eventDateRowPane.find('#monthSize-cell').onChanged(value => {
			draftDom.findAll('.crater-event-content-item-date-month').forEach(element => {
				element.css({ fontSize: value });
			});
		});

		eventDateRowPane.find('#fontFamily-cell').onChanged(value => {
			draftDom.findAll('.crater-event-content-item-date').forEach(element => {
				element.css({ fontFamily: value });
			});
		});
		//appends the dom and pane prototypes to the dom and pane when you click add new
		this.paneContent.find('.new-component').addEventListener('click', event => {
			let newEventRowDom = eventListRowDomPrototype.cloneNode(true);
			let newEventRowPane = eventListRowPanePrototype.cloneNode(true);

			eventList.append(newEventRowDom);//c
			this.paneContent.find('.list-pane').append(newEventRowPane);
			eventRowHandler(newEventRowPane, newEventRowDom);
		});

		let paneItems = this.paneContent.findAll('.crater-event-item-pane');
		paneItems.forEach((eventRow, position) => {
			eventRowHandler(eventRow, eventListRow[position]);
		});

		let showHeader = titlePane.find('#toggleTitle-cell');
		showHeader.addEventListener('change', e => {

			switch (showHeader.value.toLowerCase()) {
				case "hide":
					draftDom.findAll('.crater-event-title').forEach(element => {
						element.style.display = "none";
					});
					break;
				case "show":
					draftDom.findAll('.crater-event-title').forEach(element => {
						element.style.display = "flex";
					});
					break;
				default:
					draftDom.findAll('.crater-event-title').forEach(element => {
						element.style.display = "none";
					});
			}
		});

		//to hide or show properties
		let showIcon = eventIconRowPane.find('#toggleIcon-cell');

		showIcon.addEventListener('change', e => {

			switch (showIcon.value.toLowerCase()) {
				case "hide":
					draftDom.findAll('.crater-event-content-item-icon-image').forEach(element => {
						element.style.display = "none";
					});
					break;
				case "show":
					draftDom.findAll('.crater-event-content-item-icon-image').forEach(element => {
						element.style.display = "block";
					});
					break;
				default:
					draftDom.findAll('.crater-event-content-item-icon-image').forEach(element => {
						element.style.display = "none";
					});
			}
		});

		let showTitle = eventTitleRowPane.find('#toggleTitle-cell');
		showTitle.addEventListener('change', e => {

			switch (showTitle.value.toLowerCase()) {
				case "hide":
					draftDom.findAll('.crater-event-content-task-caption').forEach(element => {
						element.style.visibility = "hidden";
					});
					break;
				case "show":
					draftDom.findAll('.crater-event-content-task-caption').forEach(element => {
						element.style.visibility = "visible";
					});
					break;
				default:
					draftDom.findAll('.crater-event-content-task-caption').forEach(element => {
						element.style.visibility = "hidden";
					});
			}
		});

		let showLocation = eventLocationRowPane.find('#toggleLocation-cell');
		showLocation.addEventListener('change', e => {

			switch (showLocation.value.toLowerCase()) {
				case "hide":
					draftDom.findAll('.crater-event-content-task-location-place').forEach(element => {
						element.style.visibility = "hidden";
						element.previousSibling.style.visibility = "hidden";
					});
					break;
				case "show":
					draftDom.findAll('.crater-event-content-task-location-place').forEach(element => {
						element.style.visibility = "visible";
						element.previousSibling.style.visibility = "visible";
					});
					break;
				default:
					draftDom.findAll('.crater-event-content-task-location-place').forEach(element => {
						element.style.visibility = "hidden";
						element.previousSibling.style.visibility = "hidden";
					});
			}
		});

		let showDuration = eventDurationRowPane.find('#toggleDuration-cell');
		showDuration.addEventListener('change', e => {

			switch (showDuration.value.toLowerCase()) {
				case "hide":
					draftDom.findAll('.crater-event-content-task-location-duration').forEach(element => {
						element.style.visibility = "hidden";
						element.previousSibling.style.visibility = "hidden";
					});
					break;
				case "show":
					draftDom.findAll('.crater-event-content-task-location-duration').forEach(element => {
						element.style.visibility = "visible";
						element.previousSibling.style.visibility = "visible";
					});
					break;
				default:
					draftDom.findAll('.crater-event-content-task-location-duration').forEach(element => {
						element.style.visibility = "hidden";
						element.previousSibling.style.visibility = "hidden";
					});
			}
		});

		let showDate = eventDateRowPane.find('#toggleDate-cell');
		showDate.addEventListener('change', e => {

			switch (showDate.value.toLowerCase()) {
				case "hide":
					draftDom.findAll('.crater-event-content-item-date').forEach(element => {
						element.style.display = "none";
					});
					break;
				case "show":
					draftDom.findAll('.crater-event-content-item-date').forEach(element => {
						element.style.display = "block";
					});
					break;
				default:
					draftDom.findAll('.crater-event-content-item-date').forEach(element => {
						element.style.display = "none";
					});
			}
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = draftDom.innerHTML;//upate the webpart
			this.element.css(draftDom.css());
			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;
		});
	}

	public update(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;
		this.paneContent = this.setUpPaneContent(params);

		let paneConnection = this.sharePoint.app.find('.crater-property-connection');

		let updateWindow = this.elementModifier.createForm({
			title: 'Setup Meta Data', attributes: { id: 'meta-data-form', class: 'form' },
			contents: {
				title: { element: 'select', attributes: { id: 'meta-data-title', name: 'Title' }, options: params.options },
				icon: { element: 'select', attributes: { id: 'meta-data-icon', name: 'Icon' }, options: params.options },
				day: { element: 'select', attributes: { id: 'meta-data-day', name: 'Day' }, options: params.options },
				month: { element: 'select', attributes: { id: 'meta-data-month', name: 'Month' }, options: params.options },
				location: { element: 'select', attributes: { id: 'meta-data-location', name: 'Location' }, options: params.options },
				start: { element: 'select', attributes: { id: 'meta-data-start', name: 'Start' }, options: params.options },
				end: { element: 'select', attributes: { id: 'meta-data-end', name: 'End' }, options: params.options }
			},
			buttons: {
				submit: { element: 'button', attributes: { id: 'update-element', class: 'btn' }, text: 'Update' },
			}
		});

		let data: any = {};
		let source: any;
		updateWindow.find('#update-element').addEventListener('click', event => {
			event.preventDefault();
			data.title = updateWindow.find('#meta-data-title').value;
			data.icon = updateWindow.find('#meta-data-icon').value;
			data.day = updateWindow.find('#meta-data-day').value;
			data.month = updateWindow.find('#meta-data-month').value;
			data.location = updateWindow.find('#meta-data-location').value;
			data.start = updateWindow.find('#meta-data-start').value;
			data.end = updateWindow.find('#meta-data-end').value;
			source = func.extractFromJsonArray(data, params.source);

			let newContent = this.render({ source });
			draftDom.find('.crater-event-content').innerHTML = newContent.find('.crater-event-content').innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = draftDom.outerHTML;
			this.paneContent.find('.list-pane').innerHTML = this.generatePaneContent({ list: newContent.findAll('.crater-event-content-item') }).innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
		});


		if (!func.isnull(paneConnection)) {
			paneConnection.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
				this.element.innerHTML = draftDom.innerHTML;

				this.element.css(draftDom.css());

				this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart
			});
		}

		return updateWindow;
	}
}

class Power extends CraterWebParts {
	public key: any;
	public element: any;
	public paneContent: any;
	public elementModifier: any = new ElementModifier();
	public params: any;
	public counter: number = 15;
	public image = {
		loading: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAIAAAACACAYAAADDPmHLAAAABmJLR0QA/wD/AP+gvaeTAAAIG0lEQVR4nO2dfYgVVRiHn/zITc1V0VXXwrIyUNIgQiopQYPM7MNWMoqNoiyCPoRKrUgtiyATLAtC/5KIsCAsC0z63hDCz0ApadXUEk1dV3NF07U/zr145+zcuWfuzJw5c+d94MDO8s573pnzu2fOeefMDAiCIAiCIAiCIAiCkBcuSMDnhcC9wB3AeGAI0BM4DGwHfgbWABuBcwnUL6TI/cDfqIatVPYA7wKTgB5pBCvERz3wCWYN71eOACuB6UBvy7ELERkMbKX6xtdLB7AaeBgYZPE4hCroCfyEfyO+jRoD9AV6ASOBZlRPcdxnH79yBvgOeAYYYemYhBC8QNdG2wxcVWG/OmAasAI46OOjXNkEvAKMjfk4hCroC7ThbaDvgYtD+ukO3AwsAXZiLoZWYDEwoeBDsMwjeBukDRgeg99rgfmonsRUDAeA5cBUVO8iWOBTvI3wVgJ1XA7MBn5AjQdMxHAMWIWaktYnEJNQoBXviR+fcH2DUb3O58BJzMRwClgLPAE0Jhxf7ujAe7IHWqz7ItQg8gNU9296qdgGLABGW4y1ZtFPbhJpZRN6ArcC7wH7fOIqV7YDbwDXk17smUY/oa4wBvUr34C5GA6iMpHTUPcyBANcFUApI1FJpBbgLGZiOIIaRDYD/eyHnB2yIIBSBqMa9QvU4NBEDCeBdSgRDbUfsttkTQCl9EF19yuBdszEcBbVk8wBRtkP2T2yLIBS6oDJwFLMb2UXZxRvojKRuRxE1ooASukGXIcaRP6GuRh2owQ0mRytbahFAeiMQXX5LUAnZmI4hLq0zEBdamqWPAiglBHALNQg8jRmYugo2DcD/e2HnCx5E0ApA1GNuopwaxtaUDOKS+yHHD95FkApuU1LiwC60h01K1gK7MVcDK2FfTI1oxABVKam09IigHBcjZpRrMc8Ld0GfAg04eBqaRFA9TSi1iisxTwtfRRYRuX1ltYQAcRDPWr10irUaqZKQjgFvI4afKaKCCB+emGelt4CXJpOmAoRQLIUZxSLgb/wF8GfpCgCEYA9egAzUc9T6ud9KymtWxAB2Kcv8Bldz/3yNIIRAaRDN+BjvOe+E7jJdiAigPSoA37Fe/6/sh2ECCBdJuA9/2eBK20GIAJInxa8bTA7rINucUckWOUjbfsWm5VLDxCOqfg/uLIXmFKlz7Garx3RwzRHBBCOoNvDe6r02V/z0xbWgVwCzGgAFqEeWW9DLfXaCLwEDEgxrnZt2+rT0HnpAZoIvkmzH5ho4GcK/r3AHuC2CPGl1g55EEATZvfuO4AbUopRBJAQDZg/NXQO9XqbXinEKQJIiNfwHt8p1EuxGgtlDl0Xc8xKIU4RQEJswXt8c3xs5mo2a61Fdx4RQELo6/2H+dgM02wOWIvuPJHaIcoSZL2yzCxnNsT0+ErtOrH/yrpI7VCreYDSefu/hbIJdV1vSLDeoPOZRCYwVVy9BMwgeN5+DDW9q4Tp8ZnaJZEJDFN/7LgogBmYPcXbSWURiABcrbgMDZgtqy6WdtRrY8oRtwAkE5gwi/DGYzJvfzXAX9wCSAoRQIFq5u2bAvyJAFytuAzVzNuPBfjLhQBqKQ9gGk9adkkRqf4oeYB9JX/vjeAH0pu3CxEojmqjjmJdnbfn4hKQNi7P2+O2SyoTmFkBuD5vj9vOyURQmvcCnsb7faHTqKnb8EKZW/hfkX7AU9aiExLH9Xl73HY1lQmMY9Sud/9+n3Rp1GxszttlEFiGvIzaRQA+5GnULgLQyNuoPRcCCDMLkFF7zsnbqD0XPUAY8jZqj9su85lA10+w63aSCRSyjeu/MNftnMwEhlk8oDt3bcGF63ZJEal+uQTkHBFAzhEB5BwRQM4RAeQcEYA9Mv90sOvzbNftJBMoZBvXf2Gu20kmMOd2SSGZQKF6wgigM7EohNQII4Aj2rbfgpDh2rb+MmPBMcIIYIu23exjo/+vNVw4gss8jne0eQq1BrAR9cufR9dXsCwI8Of6qD1uu6SwVn8dsMunwnKlDRgU4M/1BovbLvNrAgFuRL0avVLjnwXuqeDL9QaL287JTGA1TCT4w8ZHqdz4+OxX63Y1IwBQ36p5EdgAnEDNEDYACwnu9ktxvcHitst8JjBu9GBdy9zFbZcUkepPMxN4XNs2ySsEPWgiVEGaAtBzBJJXCM+F2vZpXytH8fskS6W8wsIAf66PAZLgMq3uXZbrj0TYx82Pku28QhI8qNX9o+X6I9OE+QsnplfwlUcBfK3VPd9y/bHQRPDn2Y5SufHx2c8Vu0nA78BuYDlwJ9A7wJ8pt/vUPS4Gv6kwGPXq9o2o2cHxwt+1kFfY6WNzAlgNPAoMCfBdjlHAIc3nuir81BSuCiAoa3oOlTZfjxr0jgmop8hUujb+GTL8648LkxdYDNdsgtYrxCWAu4F/fOzKlT+AJajM4TDUB6FHAQ8B35TZ57mA+HLDZrwnZa6PzTzNxtYrbHoCk4F3CHdH1aS8HxBbrnA9r1DKOOBl4BfMZkF+5T/gWcP6coHreYVyNKIW23wJnDSMfQ1wTRV11Twu5xVM6FOIaxnq7up+lFB3AN+i7sCaDBRzTVx5BV1I5YhTAEJMxJFXOIy3Yf32G6TZnIgUteAULXgb9zEfG32h7DZr0QmJ8zzexm0H7gN6FMpMul5qFqcSqZAI9XTNuBW7+RM+/+8ARqQSqZAYd6FStSZTsydTilFImAcIXgp/GknI1DxXACtQq5+LDX8IWAmMTjEuIQWGAgOQR+sFQRAEQRAEQRCE7PE/g/hP40wSVzUAAAAASUVORK5CYII=',
		loading2: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAIAAAACACAYAAADDPmHLAAAABmJLR0QA/wD/AP+gvaeTAAAQbklEQVR4nO2de3Bc1X3HP79zV7uy1vJDMsKYtwPGDG1I0lAoAwlQkqEBQ5oBMm0I2KK8bWpjOclMM42m7bSh2EAsu5MXyGX6mGJaZKwSnkNSUlpCAskkmNQPPSzjl2w9d/XYvff++sfKWLsryfu4q9Wu7mdm/7hX2nN+e+93z/md7znnLvj4+Pj4+Pj4+Pj4+Pj4+MwWxOsCGxsbA9XHzrtF0RUClymcLkoQ4ZjALld4S3D/c33T3b/yum6f7PFUABvXNH9JlCeBszP49zaEFoGWsw6H37p9++2Ol7H4ZIYnAnj0a09VB4bMd4E/zbGIbuAFVFsGo/Ja47ZVI17E5XNq8hbA4+t+WKNx61XgUx7EA0gE9EcCLRLgxXVPrurzplyfichLAI2NjYF5x859VeGa8ecDAVi+3OaCpQ41C8FYSiRiOHxE6Og0HOgy2E5GVceBN1BtsUV3fH3L3QfzidcnnbwE8PiabetVdeP4czULlc9/Ls7Che6k74vb0LXfor3D0LnfMDqaURgq8DMXWgKW+/za79z9f/nE7pMgZwE81vBM2Iw4XcDCE+fOOEP5wg2jBIOZl+O6cPCQob3D0NEeIBLN9J36gSA7EPP8uqY73xFEs/oAPkAeAti0pnklSvOJ42AQvnzbKHPn5nEfFI52G9o7LNo7hN5ek+k7P1TYgdJSHY/9+L7v3xfPPYjZRc4CePyhbdtV9NYTx5d+3OHKP/D2uvf1C+3tFu3thqPdBs1MW70oL6rRFg0FfrRh450ZtymzkdxbgNXNe4GPnTj+0h/HOL1u8n4/X6JRobPT0NZuOHjQwsmsqhFVXhVDSwW68+Gm+u6CBVii5COAKFB14rj+rlFCldPTDcdtOPihYV9bIpGMxTL6GC7wPyg7wTy/futduwscZkmQjwCS7vYD944UwFg+NbYNHx5MCKGjwzA8nFkQCr8UaEHcltlsS3sngPtOmncSrCRQPRcTcBCxAVC1UDuAMxLHGR4A1/vWQhWOHTd0dhj27LXo68/443WiusMYdlaNxn8ym5JIbwVgLIK1C7GsfmCqGxzAdcI4ow52ZBDcwkwD9PYYOvYbOjoNh4+YqUMaQ+G4EV5UZadbab1Y7kmkZwJ4cLUSqg0hDGVZksHVMG7MYA9GUbsw0wCRiLC/y6Kj09DVZXAzSyKHQV5H2G4sfaEcbWnPBLDuWwYhlnc4ShgnFsCORtHR4TzLm5jRUejotOjsNOzvsohn1uA7wP8qbHfE/vevN91zoCDBTTOeCeCRb9n5R5OCEsJ1KnGiMZyhCBm14VliO8KHBxJzFO0dVsZJJLAL2K6izzY01e/yPLBpYkYLYDxKENeZgzvkYA8NgHrvOajC4cMyNry0iEQyvjxtqLaq6vZIXddbjY2NhTNEPKZkBJCMhathnFFwIoOoXZikvbfHsK8tkUR2H8vYlu4W4SVg+0Bt+OXGxtvz7RcLSokKYDyCq3MTSWQ0isYKk0QODAodHQnz6ciRjG3pKMgbCNsrCLQ83HTHQEGCy4MyEEAyyhzceAhneGQsb/CekWHo7LLY12Zx4IDByWwUO6LIT0FbrUDg2XVPfvVQQYLLkrITwHiUEK49B2c4hjM8WBDzybbhw9xs6feAVsty/7WYaxvKWgDJBHDdKpwY2AP9BTGfHAcOHTYfdRVDQ1mMKISd6rit6/+h/r+nc23DLBLAeE6aT85gBNce9byG8bb03n0WvX2Z29KCvoxIa3g09lKhbelZKoDxCEoVbryCeGQIHc3WycyMnh6hrcOivd3i2LGML3s3yk4xtAwM8mohVkv7Akgh2XwaLEgdg4MyNntpcfBQZiMKheMi/KOr1pYNW+5s9yoWXwBTUoHrzsEednEihTGfRoahoyvRMhzokkxWS48A3x6M8KgXLYIvgIwZZz4NDqCO95832ZYOMDzFVIjA28ayb1n7nXuO5FOnL4CcGJu0GrWwIxE07n0S6SocPGhob7fY1zbpQpd2G/eqfPZLZOxv+oxHESIEQv1U1jpULq4mWLsIq2quZzUYgbPOdLn6qjhfvWOU666JU1WVliycH0Batz64NeeKfQF4gDCMVdFHcN4IlYvnEjytFqt6fuIueoBl4KKLHP7ky6Occ05qHiKfHDZVj+Vati8AjxFGsKx+guEoc+pChOpqCMyrAWPlXXYwCF+4IcbS85NFIHDvpjXNV+RSpi+AgmJjzAAVVQPMqQsQOr2GioW1mEAo5xJF4A+vi1GzMKk7MKJ8M5fyfAFMGw5GBgiE+gktcqlcvICK2kVIZdWp35pCIABXX51sECr80RMPPL0027J8ARSFsSSyoo/KBbFxSWR1xiUsOcNN3YhjHEtWZBuJL4AZwMkkcpjKxVWE6mqwqhecMolctix5QktStulnQiDbN/jkRtse4bVWQ2Qg+aZWz4PrV7icf0Hi2yzEEBMjGAbCIVw3jDPsYEf706azz1iclgxekm1cfgswTbzeaqXdfIDBAXitdbLbYGNMPxXhCJV1c6ioWcT4W1ad0mMonJZtXH4LkAFDUeG9t4W2PULf8cRNXFirnH+h8snLlapw4afvhRiBYAxrcRXxAXCGIgQr0uqdn225vgBOwe5dwis7LGIpSzuPHhaOHhbeexs+/0WHZRdPLYLrV7i8ttMwmLIqsHo+fO6mzBenCCME51nErAU4g2n7VLJ2nnwBTMHuXULrc9aU2xFiMWjdbnHTbVOL4PwLXO5Z59VsokMwPEzMmQd5bsbxc4BJGIomvvkZ7UVReGWHxXBh1pJMgkPFvPyXtfkCmIT33pa0Zl9rHNxzR3HPHUVrkmc/Y6Pw3tvTezmF/GchfQFMQtvu5EujNTa6wAYLsEAXOGkiaNtdhAck5IkvgEno600+1ur0/jv1XF+PL4CyIZ6aW1kTJAMp51K7jFKgLEcBM2HcnkqmTuB0U3YC8Grc7jUJJzD9/Akn8J61xRFAWXUBJ8btUzXFJ8btuz8ovf66EJSNAGb6uP36FS7V89LPZ+sEek3ZdAGTjdu1OjFUk0EL6Tn5cU+M26+8dnqaXm+dQO8omxZgtozbvaZsBDBbxu1eUzYCmC3jdq/JJwc4CCwBmDsvvyHVTBy3zxbyEIA8APqDqjB119+Ye3IzU8fts4WcBbB+y8oXgBe0+ynN9fl9Xs63z3RmqhPoQQ6QWyI108ftXpPbmsDCU7SaS2G+fTZQtCs628btZeUEPnb/M3UmYD+syI1Nf5soJtusvT9t3J5+EbTaSXLvSnncPlOdwKwFsGl1822o8xRItcBHT9rONmtPG4NPtHk25Zw/bveerLqATaubbwP+DWHSTWz+bFtpkbEAHrv/mTqUp8gk7S+DrH22kLEATMB+OPWb72ftpU8Wd0huGn9U7ln7bCGLJFA+Nt61mWy2TXpOHpdy1u41ZeAEavKTqDKcbXPN2Si5PxKlXJipTmDBVwQ5ZhmYZQhRjHsU0W5EC/MIVp/smTbpKWEccz629fvY1pXTVe2MoaycwHxRmVOMaovKTHUC/XHaLMcXwCyn5ASghIsdQllRcvsC7MAViA4j2o3Ro4gOUIhfFJ0tlJwAIJFEqpyDyzkII4geB/YWO6ySpCQFMB6lEpUzmekCKAMnMDfc2PH/wvVn8meqE1jwmitPv+yzwdrlISd64Hp3pLs18chLn5nCtElvzlmfeb3yjMtXBE/73QXx4Z7fm656T830TFj5TuA4wks+/e6m1c3FqDqNeOBqjNs7NqroBgpzM2aqE1jySWD+VOCaOqAOR12EPoweA2bEbzsXnJIzggqKGFRqEjOYs4TMBaAkz+FO1FLaaf3pBE/FyRNnyF9p6CEZC0CFfeOPZTB9HbdEUovTfWn/lCfBRb8Tdob3f8YdOvg8dqTPdwHzI/McQKUV0U+cODyxYUOrHVBBIiZpE0fiLez0KM4k5iy55k3gTYBI25uXVoRkA/CVQtRV7mQsAHVMk1jOn49fGSw9gbSbPo5+RuNN+QZ4KuYuvfpXwB2bVjfPCAHYgSsQtxujxxA9aXmUvBO44bt3HlWj9WS4n1eF+obv33cs99BKEyWMa87Dtj6NHbgKx7oINbXl4QQ2NNU/p6K3M3Vy16/CrQ1Nq/4jv9BKHyWEK2dhm08UZGGsSk3eZWQtvYam+ucqRC8A/hp4FySSePGuiv6VxmIXlNTN1+lJIq/94jLmzksXQfV8yckJdE0dtvXxvOPKyQh6uKm+G/jLsVdJ40b3XAihtRII30JwwVliCuONnXdhDas2TPzrrqJDOHoMo91jecPkolRCuGYprlniSVzFcwKVwaStZg7pO4SnwVeoPPuGvcBqYLW+/1JNbH7oQTWhrwDLva5rMlSqPlrfAHGMHke0F9EoiQtTMfY/tbhSw8RbqXOjaNnHTPEVkuq75Iae0FnX/k3lkisvLmQ9U1OBK4txzMWJRNK6HNv6FI5ZjiunMf7mO05aS5H1tHvx0k+V1vGH0hNA+qyE4G1B+qxp8xVKlaHBtJ+MyXoCo3gtgGOaUu1l6QlgOkOY/cGJ/IVp8RVKiUP7U3pEka5syyiaAHxfIX9++8sjySdU38i2jKLOBpajr2D3/nqNRrve0Xh/HOBAex//vPnn/NPmd/jxzj107unBsfN3/fbv7aVzT0/SOaPakm05RV8P0NBU/9zmNU//JK6yBrgRZGwuVneraCuj8aZS+uZXLb1lC7BFtdHEOi69+eVnI/8yFLHnAPR2D/Hrnx0kGLQ4Z1kNS5fXcu6yGirnVGRVx0DPMK8890HySdU31m2tfzfbeIsuACgvX+EEIo0u0LJpdXM/kLQZMhZz2Pubbvb+phsxwpnnzue8i2pZuryW+bVT75vs2tfLK899wHA0Pv60q5bZkEucM0IAnjBDfIVUVHSNqPwAWDDh313lQHsfB9r7+OlL+6g5LczS5TWcfWENtXVhQpUBhiIxDu3v57fvHaEjpdlPFKLfbNi86he5xFc2AlBhn8DJ6epBC12QbLFOt68AiS6ucWVza3WY64CbEVYw9pT1iejpjtLTHeXnb2aY0Iv8cP2WVX+Xa3xlI4CZtF4hlcZtq0aAF8de9z/+0LZLFL1pTAxXktvSZAflLx7Zctffr2dlzrGVjQBm6nqFiXhk68r3gfeBR59Y23ye68jNqqwQ9LPAKTNCRV7D8I2GzSt/sZ5VecVSVk9x2rjm6VtF5VlO/bn0VEPLTaubk/wJd+nEP9Rs2pJn+NZvWZXzNX1ibfMCjXODCteLcJkqZwNzBTmiaKeKvm657Mgl25+MshIAfCSCp4AJtmEACV+h/lS+QjEEUAzKblm4h+sVklYfizPBfU2fxi+5FctlkwOMxxtfQTtBPpoV1KhA6m8jRdNmMDtyr684lF0L4BkiL48/ND0BJHLyF04kajA9KQIQXpqu8LzCF8AkuGptBk52/K4gRwNIZwjpCCFHKsBN6hZG1XU2T3ec+eILYBI2bLmzHZGvpZ4XN/FKOy/a0LD1zzqnIzYv8QUwBeubVm4WkQam3DIsNsr6R5rqt0xbYB5SUkOWYvH4Q9suUeEboDcCC8dO94rQ6qLfbmiq31XM+HymkY33fm/Rxnu/t6jYcfj4+Pj4+Pj4+Pj4+Pj45MD/A+JWzX0rH59IAAAAAElFTkSuQmCC'
	};

	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.params = params;
	}

	public render(params): any {
		let power: any = this.createKeyedElement({
			element: 'div', attributes: { class: 'crater-power crater-component', id: 'crater-power', 'data-type': 'power' }, children: [
				{
					element: 'div', attributes: { id: 'power-overlay' }, children: [
						{ element: 'div', attributes: { id: 'power-overlay-text' } }
					]
				},
				{
					element: 'div', attributes: { class: 'crater-power-container' }, children: [
						{
							element: 'div', attributes: { id: 'renderContainer' }, children: [
								{
									element: 'div', attributes: { id: 'render-error' }, children: [
										{ element: 'p', attributes: { id: 'render-text' } }
									]
								}
							]
						},
						{
							element: 'div', attributes: { class: 'crater-power-timer' }, children: [
								{ element: 'span', attributes: { class: 'crater-power-counter' }, text: this.counter + ' Seconds to Login!' }
							]
						},
						{
							element: 'div', attributes: { id: 'crater-power-connect', class: 'crater-power-connect' }, children: [
								{
									element: 'div', attributes: { class: 'user' }, children: [
										{ element: 'img', attributes: { class: 'crater-power-image', alt: 'Master-User', src: this.image.loading } },
										{
											element: 'div', attributes: { class: 'login-container' }, children: [
												{ element: 'h4', attributes: { class: 'user-header' }, text: 'Master Account' },
												{ element: 'p', attributes: { class: 'power-text' }, text: 'SharePoint site visitors will not be required to login to Power BI to view Power BI Data in SharePoint.' },
												{ element: 'p', attributes: { class: 'user-recommended' }, text: 'Good for 20+ users' },
												{ element: 'button', attributes: { id: 'master', class: 'user-button' }, text: 'Connect' }
											]
										}
									]
								},
								{
									element: 'div', attributes: { class: 'user' }, children: [
										{ element: 'img', attributes: { class: 'crater-power-image', alt: 'Logged-in-User', src: this.image.loading } },
										{
											element: 'div', attributes: { class: 'login-container' }, children: [
												{ element: 'h4', attributes: { class: 'user-header' }, text: 'Logged-In User' },
												{ element: 'p', attributes: { class: 'power-text' }, text: 'Every SharePoint Site Visitor will be required to login to Power BI using Pro Account to view Power BI Data' },
												{ element: 'p', attributes: { class: 'user-recommended' }, text: 'Good for 1 - 20 users' },
												{ element: 'button', attributes: { id: 'normal', class: 'user-button' }, text: 'Connect' }
											]
										}
									]
								}
							]
						},
						{ element: 'div', attributes: { class: 'login_form' } },
					]
				},
			]
		});

		let form = `<div id="id01" class="modal">
		  <form class="modal-content animate-modal">
		 	 <div class="form-header">
		  		<div class="form-header-space">
					<h4>MASTER ACCOUNT LOGIN</h4> 	
			  	</div>
			</div>  
			<div class="form-body">
				<div class="form-container">
			  		<label for="uname"><b>Username: </b></label>
			  		<input class="input" id="power-username" type="text" placeholder="Enter Username" name="uname">
				</div>
				<div class="form-container">
			  		<label for="psw"><b>Password: </b></label>
			  		<input class="input" id="power-password" type="password" placeholder="Enter Password" name="psw">
				</div>
				<div class="error-message" id="emptyField">
					<p> The above field(s) cannot be empty! && Name length cannot be less than 3</p>
				</div>
			</div>
			<div class="cancel-container">
			  <button id="login-submit" class="cancelbtn" type="submit">Login</button>
			  <button type="button" id="cancelbtn" class="cancelbtn">Cancel</button>
			</div>
		  </form>
		</div>
		`;

		power.find('.login_form').innerHTML += form;

		this.key = power.dataset.key;
		this.sharePoint.properties.pane.content[this.key].settings.myPowerBi = { showNavContent: '', showFilter: '', loginType: '', code: '', username: '', tenantID: "90fa49f0-0a6c-4170-abed-92ed96ba67ca", clientSecret: 'FUq.Y0@BN4byWh6B8.H:et:?F/VX2-3a', password: '', clientId: '9605a407-7c23-4dc8-bd90-997fbc254d38', accessToken: '', embedToken: '', embedUrl: '', reports: [], reportName: [], groups: [], groupName: [], dashboards: [], dashBoardName: [], tiles: [], tileName: [], width: '100%', height: '300px' };
		window.onerror = (msg, url, lineNumber, columnNumber, error) => {
			console.log(msg, url, lineNumber, columnNumber, error);
		};
		power.find('#master').addEventListener('click', event => {
			let loginForm = power.find('.login_form') as any;
			let powerConnect = power.find('.crater-power-connect') as any;

			powerConnect.style.display = "none";
			loginForm.style.display = 'block';
			let username = power.find('#power-username') as any;
			let password = power.find('#power-password') as any;
			let errorDisplay = power.find('#emptyField') as any;
			let loginButton = power.find('#login-submit') as any;
			let cancelButton = power.find('#cancelbtn') as any;
			let renderBox = power.find('#renderContainer') as any;

			loginButton.addEventListener('click', e => {
				e.preventDefault();
				if ((username.value.length == 0) || (password.value.length == 0) || (username.value.length < 3)) {
					username.style.border = '1px solid tomato';
					password.style.border = '1px solid tomato';
					errorDisplay.style.display = "block";
					return;
				}
				else {
					this.showLoading();
					this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.accessToken = '';
					this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.username = username.value;
					this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.password = password.value;
					this.getAccessToken().then(aResponse => {
						if (this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.accessToken) {
							this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.groups.length = 0;
							this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.groupName.length = 0;
							this.getWorkSpace(aResponse);
							renderBox.find('#render-error').style.display = 'none';
							renderBox.style.display = "block";
							this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.loginType = 'master';
							loginForm.style.display = "none";
							if (renderBox.find('.connected')) renderBox.find('.connected').remove();
							renderBox.makeElement({
								element: 'div', attributes: { class: 'connected' }, children: [
									{ element: 'img', attributes: { alt: 'Master-User', src: this.image.loading2 } },
									{
										element: 'div', attributes: { class: 'login-container' }, children: [
											{ element: 'h4', text: 'Power BI Master' },
											{ element: 'button', attributes: { id: 'master', class: 'user-button' }, text: 'Reconnect' },
											{ element: 'p', text: `Connected using ${this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.username}` },
										]
									}
								]
							});
							renderBox.find('#master').addEventListener('click', ev => {
								loginButton.innerHTML = 'Login';
								if (power.find('.login_form').style.display = 'none') power.find('.login_form').style.display = 'block';
							});


						}
					});
				}
			});

			cancelButton.addEventListener('click', e => {
				loginForm.style.display = 'none';
				if (!renderBox.find('.connected')) powerConnect.style.display = 'grid';
			});

		});
		power.find('#normal').addEventListener('click', event => {
			// this.checkConsent();
			this.startCounter();
			this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.accessToken = '';
			this.requestCode().then(response => {
				setTimeout(() => {
					this.getAccessToken(response).then(access => {
						this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.groups.length = 0;
						this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.groupName.length = 0;
						this.getWorkSpace(access);
						if (this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.accessToken && this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.groups.length !== 0) {
							let powerConnect = power.find('.crater-power-connect') as any;
							powerConnect.style.display = "none";
							let renderBox = power.find('#renderContainer') as any;
							renderBox.find('#render-error').style.display = 'none';
							renderBox.style.display = 'block';
							power.find('.crater-power-timer').style.display = "none";
							if (renderBox.find('.connected')) renderBox.find('.connected').remove();
							renderBox.makeElement({
								element: 'div', attributes: { class: 'connected' }, children: [
									{ element: 'img', attributes: { alt: 'Login', src: this.image.loading2 } },
									{
										element: 'div', attributes: { class: 'login-container' }, children: [
											{ element: 'h4', text: 'Power BI User' },
											{ element: 'p', attributes: { id: 'new-user', class: 'user-button' }, text: 'Connected' },
										]
									}
								]
							});
							this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.loginType = 'user';
						}
					});
				}, 15000);
			});
		});

		return power;
	}

	public showLoading() {
		let loadingButton = this.element.find('#login-submit') as any;
		loadingButton.style.zIndex = 2;
		loadingButton.render({
			element: 'img', attributes: { class: 'crater-icon', src: this.sharePoint.images.loading, style: { width: '20px', height: '20px' } }
		});
	}
	public checkConsent() {
		let draftPower = this.sharePoint.properties.pane.content[this.key].settings.myPowerBi;

		let promise = new Promise((res, rej) => {
			let host = (location.href === `https://localhost:4321/temp/workbench.html`) ? location.href : (location.origin === `https://ipigroup.sharepoint.com/`) ? location.origin : `https://ipigroup.sharepoint.com/`;
			let openURL = `https://login.microsoftonline.com/${draftPower.tenantID}/v2.0/adminconsent?client_id=${draftPower.clientId}&state=12345&redirect_uri=${host}&scope=https://analysis.windows.net/powerbi/api/Workspace.ReadWrite.All`;

			let newWindow = window.open(openURL, '_blank');
			const windowLocation = newWindow.location.href;
			const changedURL = (newWindow.location.href === windowLocation) ? res(newWindow.location.href) : rej(new Error('Sorry, permissions were not granted!'));
			// setTimeout(() => {
			// }, 5000);
		}).catch(err => console.log(err.message));

		return promise;
	}
	public startCounter() {
		//@ts-ignore
		this.element.find('.crater-power-timer').style.display = 'block';
		clearInterval(1);
		setInterval(() => {
			if (this.counter === 0) {
				//@ts-ignore
				this.element.find('.crater-power-counter').render({
					element: 'img', attributes: { class: 'crater-icon', src: this.sharePoint.images.loading, style: { width: '20px', height: '20px' } }
				});
			}
			else {
				this.counter--;
				//@ts-ignore
				this.element.find('.crater-power-counter').textContent = 'Please wait ' + this.counter + ' Seconds';
			}
		}, 1000);

	}
	public requestCode() {
		let promise = new Promise((res, rej) => {
			let draftPower = this.sharePoint.properties.pane.content[this.key].settings.myPowerBi;
			let host = (location.href === `https://localhost:4321/temp/workbench.html`) ? location.href : (location.origin === `https://ipigroup.sharepoint.com/`) ? location.origin : `https://ipigroup.sharepoint.com/`;
			let state = 12345;
			let openURL = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${draftPower.clientId}&response_type=code&redirect_uri=${host}&response_mode=query&scope=openid&state=${state}`;
			let newWindow = window.open(openURL, '_blank');
			newWindow.focus();

			setTimeout(() => {
				//@ts-ignore
				if (newWindow.location.search.substring(1).indexOf('code') !== -1) {
					let splitURL = newWindow.location.search.substring(1).split('=')[1].split('&state')[0];
					draftPower.code = splitURL;
					res(splitURL);
					newWindow.close();
				} else {
					rej(new Error('Sorry, there was an error'));
					newWindow.close();
				}
			}, 13000);
		}).catch(error => {
			console.log(error.message);
			//@ts-ignore
			this.element.find('#render-error').style.display = 'block';
			this.element.find('#render-text').textContent = error.message;
		});

		return promise;
	}
	public createWorkSpace(access) {
		let draftPower = this.sharePoint.properties.pane.content[this.key].settings.myPowerBi;
		let url = `https://api.powerbi.com/v1.0/myorg/groups`;

		let accessString = 'Bearer ' + access;

		let promise = new Promise((res, rej) => {
			let result;
			let request = new XMLHttpRequest();
			request.onreadystatechange = function (e) {
				if (this.readyState == 4 && this.status == 200) {
					result = request.responseText;
					console.log('group created...');
					draftPower.groups = [];
					draftPower.groupName.length = 0;
					for (let workspace in JSON.parse(result)) {

						draftPower.groups.push({
							//@ts-ignore
							workspaceName: workspace.name,
							//@ts-ignore
							workspaceId: workspace.id
						});
						//@ts-ignore
						draftPower.groupName.push(workspace.name);
					}
				}
			};

			request.open('POST', url, false);
			request.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
			request.setRequestHeader('Authorization', accessString);
			request.send(`name=new workspace`);
			res(JSON.parse(result));
			rej(new Error('Sorry, there was an error'));
		}).catch(error => {
			console.log(error.message);
			//@ts-ignore
			this.element.find('#render-error').style.display = 'block';
			this.element.find('#render-text').textContent = error.message;
		});

		return promise;
	}
	public cloneReport(params) {
		let draftPower = this.sharePoint.properties.pane.content[this.key].settings.myPowerBi;
		let url = `https://api.powerbi.com/v1.0/myorg/reports/${params.reportID}/Clone`;

		let accessString = 'Bearer ' + params.accessToken;

		let promise = new Promise((res, rej) => {
			let result;
			let request = new XMLHttpRequest();
			request.onreadystatechange = function (e) {
				if (this.readyState == 4 && this.status == 200) {
					result = request.responseText;
					console.log('report cloned...');

					draftPower.reportId = result.id;
					draftPower.embedUrl = result.embedUrl;
				}
			};

			request.open('POST', url, false);
			request.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
			request.setRequestHeader('Authorization', accessString);
			request.send(`name=new&targetWorkspaceId=${params.groupID}`);
			res(JSON.parse(result));
			rej(new Error('Sorry, there was an error'));
		}).catch(error => {
			//@ts-ignore
			this.element.find('#render-error').style.display = 'block';
			this.element.find('#render-text').textContent = error.message;
		});

		return promise;
	}
	public cloneTile(params) {
		let draftPower = this.sharePoint.properties.pane.content[this.key].settings.myPowerBi;
		let url = `https://api.powerbi.com/v1.0/myorg/dashboards/${params.dashboardID}/tiles/${params.tileID}/Clone`;

		let accessString = 'Bearer ' + params.accessToken;

		let promise = new Promise((res, rej) => {
			let result;
			let request = new XMLHttpRequest();
			request.onreadystatechange = function (e) {
				if (this.readyState == 4 && this.status == 200) {
					result = request.responseText;
					draftPower.tileId = result.tileId;
					draftPower.reportId = result.tileId;
					draftPower.embedUrl = result.embedUrl;
				}
			};

			request.open('POST', url, false);
			request.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
			request.setRequestHeader('Authorization', accessString);
			request.send(`name=new&targetWorkspaceId=${params.groupID}&targetDashboardId=${params.dashboardID}`);
			res(JSON.parse(result));
			rej(new Error('Sorry, there was an error'));
		}).catch(error => {
			//@ts-ignore
			this.element.find('#render-error').style.display = 'block';
			this.element.find('#render-text').textContent = error.message;
		});

		return promise;
	}
	public getReports(access, groupID, third?) {

		let draftPower = this.sharePoint.properties.pane.content[this.key].settings.myPowerBi;
		let url = (third) ? `https://api.powerbi.com/v1.0/myorg/reports` : `https://api.powerbi.com/v1.0/myorg/groups/${groupID}/reports`;

		let accessString = 'Bearer ' + access;
		let promise = new Promise((res, rej) => {
			let result;
			let request = new XMLHttpRequest();
			request.onreadystatechange = function (e) {
				if (this.readyState == 4 && this.status == 200) {
					result = request.responseText;
					console.log('getting reports...');
					draftPower.reports.length = 0;
					draftPower.reportName.length = 0;
					for (let i = 0; i < JSON.parse(result).value.length; i++) {
						if (draftPower.reports.indexOf(JSON.parse(result).value[i].name) === -1) {
							//@ts-ignore
							draftPower.reports.push({
								reportName: JSON.parse(result).value[i].name,
								reportId: JSON.parse(result).value[i].id,
								embedUrl: JSON.parse(result).value[i].embedUrl
							});

							draftPower.reportName.push(JSON.parse(result).value[i].name);
						}
					}
					if (draftPower.reportName.length === 0) {
						draftPower.reportName.push('No Reports');
					}
				}
			};

			request.open('GET', url, false);
			request.setRequestHeader('Authorization', accessString);
			request.send();
			res(JSON.parse(result));
			rej(new Error('Sorry, there was an error'));
		}).catch(error => {
			console.log(error.message);
			//@ts-ignore
			this.element.find('#renderContainer').style.display = "block";
			this.element.find('#render-text').textContent = error.message;

		});

		return promise;
	}
	public getPages(params) {
		let draftPower = this.sharePoint.properties.pane.content[this.key].settings.myPowerBi;
		let url = (params.groupID === 'none') ? `https://api.powerbi.com/v1.0/myorg/reports/${params.reportID}/pages` : `https://api.powerbi.com/v1.0/myorg/groups/${params.groupID}/reports/${params.reportID}/pages`;

		let accessString = 'Bearer ' + params.accessToken;
		let promise = new Promise((res, rej) => {
			let result;
			let request = new XMLHttpRequest();
			request.onreadystatechange = function (e) {
				if (this.readyState == 4 && this.status == 200) {
					result = request.responseText;
					console.log('getting pages...');
					draftPower.pageName = [];
					draftPower.pages = [];
					for (let page in JSON.parse(result).value) {
						//@ts-ignore
						if (draftPower.pageName.indexOf(JSON.parse(result).value[page].displayName) === -1) {
							draftPower.pages.push({
								pageName: JSON.parse(result).value[page].Name,
								displayName: JSON.parse(result).value[page].displayName
							});
							//@ts-ignore
							draftPower.pageName.push(JSON.parse(result).value[page].displayName);
						}
					}
				}
			};

			request.open('GET', url, false);
			request.setRequestHeader('Authorization', accessString);
			request.send();
			res(JSON.parse(result));
			rej(new Error('Sorry, there was an error'));
		}).catch(error => {
			console.log(error.message);
			//@ts-ignore
			this.element.find('#renderContainer').style.display = "block";
			this.element.find('#render-text').textContent = error.message;
		});

		return promise;
	}
	public getDashboards(access, groupID, third?) {
		let draftPower = this.sharePoint.properties.pane.content[this.key].settings.myPowerBi;
		draftPower.dashboards.length = 0;
		draftPower.dashBoardName.length = 0;
		let url = (third) ? `https://api.powerbi.com/v1.0/myorg/dashboards/` : `https://api.powerbi.com/v1.0/myorg/groups/${groupID}/dashboards/`;

		let accessString = 'Bearer ' + access;
		let promise = new Promise((res, rej) => {
			let result;
			let request = new XMLHttpRequest();
			request.onreadystatechange = function (e) {
				if (this.readyState == 4 && this.status == 200) {
					result = request.responseText;
					console.log('getting dashboards...');
					if (draftPower.groupId !== 'none') {
						for (let dashboard of JSON.parse(result).value) {

							//@ts-ignore
							draftPower.dashboards.push({
								dashboardName: dashboard.displayName,
								dashboardId: dashboard.id,
								embedUrl: dashboard.embedUrl
							});
							draftPower.dashBoardName.push(dashboard.displayName);
						}

						if (draftPower.dashBoardName.length === 0) {
							draftPower.dashBoardName.push('No Dashboards');
						}
					} else {
						if (draftPower.dashBoardName.length === 0) {
							draftPower.dashBoardName.push('Cannot Embed Dashboards for this Workspace');
						}
					}

				}
			};
			request.open('GET', url, false);
			request.setRequestHeader('Authorization', accessString);
			request.send();
			res(JSON.parse(result));
			rej(new Error('Sorry, there was an error'));
		}).catch(error => {
			console.log(error.message);
			//@ts-ignore
			this.element.find('#renderContainer').style.display = "block";
			this.element.find('#render-text').textContent = error.message;
		});

		return promise;
	}
	public getTiles(access, groupID, dashboardID) {
		let draftPower = this.sharePoint.properties.pane.content[this.key].settings.myPowerBi;

		let url = (groupID === 'none') ? `https://api.powerbi.com/v1.0/myorg/dashboards/${dashboardID}/tiles` : `https://api.powerbi.com/v1.0/myorg/groups/${groupID}/dashboards/${dashboardID}/tiles`;

		let accessString = 'Bearer ' + access;

		let promise = new Promise((res, rej) => {
			let result;
			let request = new XMLHttpRequest();
			request.onreadystatechange = function (e) {
				if (this.readyState == 4 && this.status == 200) {
					result = request.responseText;
					console.log('getting dashboards...');
					draftPower.tiles.length = 0;
					draftPower.tileName.length = 0;
					for (let tile of JSON.parse(result).value) {
						draftPower.tiles.push({
							tileName: tile.title,
							tileId: tile.id,
							embedUrl: tile.embedUrl,
							reportId: tile.reportId
						});
						draftPower.tileName.push(tile.title);
					}
					draftPower.tileName.push('Show Full Dashboard');

				}
			};
			request.open('GET', url, false);
			request.setRequestHeader('Authorization', accessString);
			request.send();
			res(JSON.parse(result));
			rej(new Error('Sorry, there was an error'));
		}).catch(error => {
			console.log(error.message);
			//@ts-ignore
			this.element.find('#render-error').style.display = 'block';
			this.element.find('#render-text').textContent = error.message;
		});

		return promise;
	}
	public addUser(params) {
		let draftPower = this.sharePoint.properties.pane.content[this.key].settings.myPowerBi;
		let url = `https://api.powerbi.com/v1.0/myorg/groups/${params.groupID}/users`;
		const sendLink = (draftPower.loginType.toLowerCase() === 'master') ? `identifier=${draftPower.username}&groupUserAccessRight=Admin&principalType=User` : (draftPower.loginType.toLowerCase() === 'user') ? `id=${params.groupID}&groupUserAccessRight=Admin&principalType=Group` : '';

		let accessString = 'Bearer ' + params.accessToken;
		let promise = new Promise((res, rej) => {
			let result;
			let request = new XMLHttpRequest();
			request.onreadystatechange = function (e) {
				if (this.readyState == 4 && this.status == 200) {
					result = request.responseText;
					console.log('User Added');
					console.log(result);
				}
			};

			request.open('POST', url, false);
			request.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
			request.setRequestHeader('Authorization', accessString);
			request.send(sendLink);
		}).catch(error => {
			console.log(error.message);
			//@ts-ignore
			this.element.find('#render-error').style.display = 'block';
			this.element.find('#render-text').textContent = error.message;
		});

		return promise;
	}
	public addDashBoard(params) {
		let draftPower = this.sharePoint.properties.pane.content[this.key].settings.myPowerBi;
		let url = `https://api.powerbi.com/v1.0/myorg/dashboards`;

		let accessString = 'Bearer ' + params.accessToken;
		let promise = new Promise((res, rej) => {
			let result;
			let request = new XMLHttpRequest();
			request.onreadystatechange = function (e) {
				if (this.readyState == 4 && this.status == 200) {
					result = request.responseText;
					console.log('Dashboard Added');
					draftPower.tokenEmbed = 'fulldashboard';
					draftPower.dashboardId = result.id;
					draftPower.dashboardEmbedUrl = result.embedUrl;
				}
			};

			request.open('POST', url, false);
			request.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
			request.setRequestHeader('Authorization', accessString);
			request.send('name=newDashBoard');
			res(JSON.parse(result));
			rej(new Error('Sorry, the Dashboard could not be added'));
		}).catch(error => {
			console.log(error.message);
			//@ts-ignore
			this.element.find('#render-error').style.display = 'block';
			this.element.find('#render-text').textContent = error.message;
		});

		return promise;
	}
	public deleteUser(params) {
		let draftPower = this.sharePoint.properties.pane.content[this.key].settings.myPowerBi;
		let url = ` https://api.powerbi.com/v1.0/myorg/groups/${params.groupID}/users/${params.username}`;

		let accessString = 'Bearer ' + params.accessToken;
		let promise = new Promise((res, rej) => {
			let result;
			let request = new XMLHttpRequest();
			request.onreadystatechange = function (e) {
				if (this.readyState == 4 && this.status == 200) {
					result = request.responseText;
					console.log('User deleted');
				}
			};

			request.open('DELETE', url, false);
			request.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
			request.setRequestHeader('Authorization', accessString);
			request.send();
		}).catch(error => {
			console.log(error.message);
			//@ts-ignore
			this.element.find('#render-error').style.display = 'block';
			this.element.find('#render-text').textContent = error.message;
		});

		return promise;
	}

	public deleteWorkspace(params) {
		let draftPower = this.sharePoint.properties.pane.content[this.key].settings.myPowerBi;
		let url = ` https://api.powerbi.com/v1.0/myorg/groups/${params.groupID}`;

		let accessString = 'Bearer ' + params.accessToken;
		let promise = new Promise((res, rej) => {
			let result;
			let request = new XMLHttpRequest();
			request.onreadystatechange = function (e) {
				if (this.readyState == 4 && this.status == 200) {
					result = request.responseText;
					console.log('Workspace deleted');
				}
			};

			request.open('DELETE', url, false);
			request.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
			request.setRequestHeader('Authorization', accessString);
			request.send();
		}).catch(error => {
			console.log(error.message);
			//@ts-ignore
			this.element.find('#render-error').style.display = 'block';
			this.element.find('#render-text').textContent = error.message;
		});

		return promise;
	}

	public getWorkSpace(access) {
		let draftPower = this.sharePoint.properties.pane.content[this.key].settings.myPowerBi;
		let url = `https://api.powerbi.com/v1.0/myorg/groups`;

		let accessString = 'Bearer ' + access;

		let promise = new Promise((res, rej) => {
			let result;
			let request = new XMLHttpRequest();
			let self = this;
			request.onreadystatechange = function (e) {
				if (this.readyState == 4 && this.status == 200) {
					result = request.responseText;
					console.log('getting workspaces...');
					draftPower.groups.length = 0;
					draftPower.groupName.length = 0;
					for (let workspace of JSON.parse(result).value) {
						//@ts-ignore
						draftPower.groups.push({
							workspaceName: workspace.name,
							workspaceId: workspace.id
						});
						draftPower.groupName.push(workspace.name);
					}

					if (self.element.find('#getWork')) {
						let errorText = self.element.find('#render-text') as any;
						errorText.textContent = '';
						//@ts-ignore
						self.element.find('#renderContainer').style.display = "none";
						self.element.find('#getWork').remove();
					}
					draftPower.groupName.push('My WorkSpace');
				}
			};

			request.open('GET', url, false);
			request.setRequestHeader('Authorization', accessString);
			request.send();
			res(JSON.parse(result));
			rej(new Error('Sorry, there was an error'));
		}).catch(error => {
			let powDiv = this.element.find('#render-error') as any;
			//@ts-ignore
			this.element.find('#renderContainer').style.display = 'block';
			powDiv.find('#render-text').textContent = `Couldn't Fetch Workspaces!`;
			powDiv.makeElement({
				element: 'div', children: [
					{ element: 'button', attributes: { id: 'getWork', class: 'user-button' }, text: 'Retry' }
				]
			});
			powDiv.find('#getWork').addEventListener('click', even => {
				this.getWorkSpace(access);
			});
		});

		return promise;
	}

	public getAccessToken(authCode?) {
		let host = (location.href === `https://localhost:4321/temp/workbench.html`) ? location.href : (location.origin === `https://ipigroup.sharepoint.com/`) ? location.origin : `https://ipigroup.sharepoint.com/`;
		let draftPower = this.sharePoint.properties.pane.content[this.key].settings.myPowerBi;
		let endpoint = `client_id=${draftPower.clientId}&grant_type=Authorization_code&code=${authCode}&scope=https://analysis.windows.net/powerbi/api/Workspace.ReadWrite.All` + `&client_secret=${draftPower.clientSecret}&redirect_uri=${host}`;
		let endpoint2 = `grant_type=password&client_id=${draftPower.clientId}&resource=https://analysis.windows.net/powerbi/api` + `&username=${draftPower.username}&password=${draftPower.password}&scope=openid`;
		let url = (authCode) ? 'https://cors-anywhere.herokuapp.com/https://login.microsoftonline.com/common/oauth2/v2.0/token/' : 'https://cors-anywhere.herokuapp.com/https://login.microsoftonline.com/common/oauth2/token/';
		const defineEndpoint = (authCode) ? endpoint : endpoint2;

		let promise = new Promise((res, rej) => {
			let result;
			let request = new XMLHttpRequest();
			request.onreadystatechange = function (e) {
				if (this.readyState == 4 && this.status == 200) {
					result = request.responseText;
					console.log('getting access tokens..');
					draftPower.accessToken = JSON.parse(result).access_token;
					draftPower.refreshToken = JSON.parse(result).refresh_token;
				}
			};
			request.open('POST', url, false);
			request.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
			request.send(defineEndpoint);
			res(JSON.parse(result).access_token);
			rej(new Error('Sorry, there was an error'));
		}).catch(error => {
			console.log(error.message);
			//@ts-ignore
			this.element.find('#login-submit').style.zIndex = 0;
			this.element.find('#login-submit').innerHTML = 'Login';
			this.element.find('.crater-power-counter').innerHTML = "Sorry, there was an error. Please, click the connect button to try again";
			//@ts-ignore
			this.element.find('#renderContainer').style.display = "block";
			this.element.find('#render-text').textContent = `Please make sure your details are valid!`;
		});

		return promise;
	}

	public tokenListener(params) {
		let currentTime = Date.now();
		let expiration = Date.parse(params.tokenExpiration);
		let safetyInterval = params.minutesToRefresh * 60 * 1000;

		let timeout = expiration - currentTime - safetyInterval;

		if (timeout <= 0) {
			console.log('updating embed token');
			this.updateToken({ groupID: params.groupId, reportID: params.reportId });
		} else {
			console.log('report embed token will be updated in ' + timeout + 'milliseconds');
			setTimeout(() => {
				this.updateToken({ groupID: params.groupId, reportID: params.reportId });
			}, timeout);
		}
	}

	public updateToken(params) {
		let draftPower = this.sharePoint.properties.pane.content[this.key].settings.myPowerBi;

		this.getEmbedToken({ accessToken: draftPower.accessToken, groupID: draftPower.groupId, reportID: draftPower.reportId, generateUrl: draftPower.tokenEmbed }).then(response => {
			let embedContainer = this.element.find('#renderContainer') as any;
			//@ts-ignore
			let reportRefresh = powerbi.get(embedContainer);

			reportRefresh.setAccessToken(response).then(() => {
				this.tokenListener({ tokenExpiration: draftPower.expiration, minutesToRefresh: 2 });
			});
		}).catch(error => {
			//@ts-ignore
			this.element.find('#renderContainer').style.display = "block";
			this.element.find('#render-text').textContent = error.message;
		});
	}

	public getEmbedToken(params) {
		let draftPower = this.sharePoint.properties.pane.content[this.key].settings.myPowerBi;
		let tokenURL;
		if ((!func.isset(params.generateUrl)) || (params.generateUrl.toLowerCase() === 'reportname')) {
			tokenURL = 'https://api.powerbi.com/v1.0/myorg/groups/' + params.groupID + '/reports/' + params.reportID + '/GenerateToken';
		}
		else if (params.generateUrl.toLowerCase() === 'fulldashboard') {
			tokenURL = `https://api.powerbi.com/v1.0/myorg/groups/${params.groupID}/dashboards/${draftPower.dashboardId}/GenerateToken`;
		} else if (params.generateUrl.toLowerCase() === 'tile') {
			tokenURL = `https://api.powerbi.com/v1.0/myorg/groups/${params.groupID}/dashboards/${draftPower.dashboardId}/tiles/${draftPower.tileId}/GenerateToken`;
		}

		let accessString = 'Bearer ' + params.accessToken;
		let promise = new Promise((res, rej) => {
			let result;
			let request = new XMLHttpRequest();
			let self = this;
			request.onreadystatechange = function (e) {
				if (this.readyState == 4 && this.status == 200) {
					result = request.responseText;
					console.log('getting embed token...');
					draftPower.embedToken = JSON.parse(result).token;
					draftPower.expiration = JSON.parse(result).expiration;

					if (self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.changed) {
						if (self.element.find('#renderContainer')) {
							self.element.find('#renderContainer').remove();
							let powerContainer = self.element.find('.crater-power-container') as any;
							powerContainer.makeElement({
								element: 'div', attributes: { id: 'renderContainer' }, children: [
									{
										element: 'div', attributes: { id: 'render-error' }, children: [
											{ element: 'p', attributes: { id: 'render-text' } }
										]
									}
								]
							});
						}
					}
				}
			};

			request.open('POST', tokenURL, false);
			request.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
			request.setRequestHeader('Authorization', accessString);
			request.send('accessLevel=View');
			res(JSON.parse(result).token);
			rej(new Error('Sorry, there was an error'));
		}).catch(error => {
			console.log(error.message);
			//@ts-ignore
			this.element.find('#render-error').style.display = 'block';
			this.element.find('#render-text').textContent = error.message;
		});

		return promise;
	}

	public embedPower(params) {
		if (!func.isset(params.tokenType)) params.tokenType = 'embed';
		try {
			const newEmbed = (params.embedUrl) ? params.embedUrl : `https://app.powerbi.com/reportEmbed?reportId=${this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.reportId}&groupId=${this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.groupId}`;
			window['powerbi-client'] = factory;
			const models = window['powerbi-client'].models;
			const newToken = (params.tokenType === "embed") ? models.TokenType.Embed : (params.tokenType === 'aad') ? models.TokenType.Aad : models.TokenType.Embed;
			const viewMode = ((!func.isset(params.viewMode)) || (params.viewMode.toLowerCase() === 'view')) ? models.ViewMode.View : (params.viewMode.toLowerCase() === 'edit') ? models.ViewMode.Edit : models.ViewMode.View;
			const pageName = (this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.namePage) ? this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.namePage : '';
			const dashboardId = (this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.embedType === 'tile') ? this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.dashboardId : '';

			const config = {
				type: params.type,
				tokenType: newToken,
				accessToken: params.accessToken,
				embedUrl: newEmbed,
				id: this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.reportId,
				dashboardId,
				viewMode,
				pageName,
				permissions: models.Permissions.All,
				settings: {
					filterPaneEnabled: this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.showFilter || false,
					navContentPaneEnabled: this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.showNavContent || false
				}
			};

			// Get a reference to the embedded dashboard HTML element 
			const reportContainer = this.element.find('#renderContainer') as any;

			reportContainer.style.display = 'block';
			//@ts-ignore
			let report = powerbi.embed(reportContainer, config);
			this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.changed = false;
			this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.embedded = true;
			report.off('loaded');

			report.on('loaded', () => {
				this.tokenListener({ tokenExpiration: this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.expiration, minutesToRefresh: 2, reportId: this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.reportId, groupId: this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.groupId });
			});
			if (func.isset(params.switch)) report.switchMode('edit');
		}
		catch (error) {
			console.log(error.message);
			//@ts-ignore
			this.element.find('.login_form').style.display = "block";
			//@ts-ignore
			this.element.find('#render-error').style.display = 'block';
			this.element.find('#render-text').textContent = error.message;
		}
	}

	public rendered(params) {
		this.element = params.element;
		this.key = params.element.dataset.key;

		window.onerror = (msg, url, lineNumber, columnNumber, error) => {
			console.log(msg, url, lineNumber, columnNumber, error);
		};
		let draftPower = this.sharePoint.properties.pane.content[this.key].settings.myPowerBi;
		if (this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.embedded) {
			this.getEmbedToken({ accessToken: draftPower.accessToken, groupID: draftPower.groupId, reportID: draftPower.reportId, generateUrl: draftPower.tokenEmbed }).then(response => {
				this.element.find('.crater-power-container').css({ width: draftPower.width, height: draftPower.height });
				this.element.find('#renderContainer').css({ width: '100%', height: draftPower.height });
				this.embedPower({ accessToken: response, type: draftPower.embedType, embedUrl: draftPower.embedUrl });
			});
		}
	}

	public setUpPaneContent(params) {
		let key = params.element.dataset['key'];
		this.element = this.sharePoint.properties.pane.content[key].draft.dom;
		this.paneContent = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'crater-property-content' }
		}).monitor();


		if (this.sharePoint.properties.pane.content[key].draft.pane.content.length !== 0) {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[key].draft.pane.content;
		}
		else if (this.sharePoint.properties.pane.content[key].content.length !== 0) {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[key].content;
		} else {
			if ((this.sharePoint.properties.pane.content[key].settings.myPowerBi.loginType === 'master')) {

				this.paneContent.makeElement({
					element: 'div', attributes: { class: 'power-pane card' }, children: [
						this.elementModifier.createElement({
							element: 'div', attributes: { class: 'card-title' }, children: [
								this.elementModifier.createElement({
									element: 'h2', attributes: { class: 'title' }, text: 'CONNECTION'
								})
							]
						}),
						this.elementModifier.createElement({
							element: 'div', attributes: { class: 'row' }, children: [
								{
									element: 'div', attributes: { class: 'message-note' }, children: [
										{ element: 'span', attributes: { style: { color: 'green' } }, text: `POWER BI connected using ${this.sharePoint.properties.pane.content[key].settings.myPowerBi.username}` }
									]
								}]
						}),
						this.elementModifier.createElement({
							element: 'div', attributes: { class: 'power-embed row' }, children: [
								this.elementModifier.cell({
									element: 'select', name: 'WorkSpace', options: this.sharePoint.properties.pane.content[key].settings.myPowerBi.groupName
								}),
								this.elementModifier.cell({
									element: 'select', name: 'user-rights', options: ['Allow', 'Deny']
								})
							]
						})
					]
				});

				this.paneContent.makeElement({
					element: 'div', attributes: { class: 'size-pane card' }, children: [
						this.elementModifier.createElement({
							element: 'div', attributes: { class: 'card-title' }, children: [
								this.elementModifier.createElement({
									element: 'h2', attributes: { class: 'title' }, text: 'SIZE AUTO CONTROL'
								})
							]
						}),
						this.elementModifier.createElement({
							element: 'div', attributes: { class: 'row' }, children: [
								{
									element: 'div', attributes: { class: 'message-note' }, children: [
										{
											element: 'span', attributes: { style: { color: 'green' } }, text: `Please change this value if you wish to change the size of the report`
										}
									]
								}]
						}),
						this.elementModifier.createElement({
							element: 'div', attributes: { class: 'power-embed row' }, children: [
								this.elementModifier.cell({
									element: 'select', name: 'display', options: ['16:9 (1280px x 720px)', '4:3 (1000px x 750px)', 'Custom Size']
								})
							]
						})
					]
				});
			}
			else if (this.sharePoint.properties.pane.content[key].settings.myPowerBi.loginType === 'user') {

				this.paneContent.makeElement({
					element: 'div', attributes: { class: 'power-pane card' }, children: [
						this.elementModifier.createElement({
							element: 'div', attributes: { class: 'card-title' }, children: [
								this.elementModifier.createElement({
									element: 'h2', attributes: { class: 'title' }, text: 'CONNECTION'
								})
							]
						}),
						this.elementModifier.createElement({
							element: 'div', attributes: { class: 'power-embed row' }, children: [
								this.elementModifier.cell({
									element: 'select', name: 'WorkSpace', options: this.sharePoint.properties.pane.content[key].settings.myPowerBi.groupName
								}),
								this.elementModifier.cell({
									element: 'select', name: 'user-rights', options: ['Allow', 'Deny']
								})
							]
						})
					]
				});

				this.paneContent.makeElement({
					element: 'div', attributes: { class: 'size-pane card' }, children: [
						this.elementModifier.createElement({
							element: 'div', attributes: { class: 'card-title' }, children: [
								this.elementModifier.createElement({
									element: 'h2', attributes: { class: 'title' }, text: 'SIZE AUTO CONTROL'
								})
							]
						}),
						this.elementModifier.createElement({
							element: 'div', attributes: { class: 'row' }, children: [
								{
									element: 'div', attributes: { class: 'message-note' }, children: [
										{
											element: 'span', attributes: { style: { color: 'green' } }, text: `Please change this value if you wish to change the size of the report`
										}
									]
								}]
						}),
						this.elementModifier.createElement({
							element: 'div', attributes: { class: 'power-embed row' }, children: [
								this.elementModifier.cell({
									element: 'select', name: 'display', options: ['16:9 (1280px x 720px)', '4:3 (1000px x 750px)', 'Custom Size']
								})
							]
						})
					]
				});
			}
			else if (this.sharePoint.properties.pane.content[key].settings.myPowerBi.loginType === '') {
				let userList = this.sharePoint.properties.pane.content[key].draft.dom.find('.crater-power-container');
				let userListRows = userList.findAll('.user');

				this.paneContent.makeElement({
					element: 'div', attributes: { class: 'layout-pane card' }, children: [
						this.elementModifier.createElement({
							element: 'div', attributes: { class: 'card-title' }, children: [
								this.elementModifier.createElement({
									element: 'h2', attributes: { class: 'title' }, text: 'Edit Layout Box'
								})
							]
						}),
						{
							element: 'div', attributes: { class: 'row' }, children: [
								this.elementModifier.cell({
									element: 'input', name: 'backgroundcolor', value: this.element.find('.user').css()['background-color']
								})
							]
						}
					]
				});

				this.paneContent.makeElement({
					element: 'div', attributes: { class: 'layout-button-row-pane card' }, children: [
						this.elementModifier.createElement({
							element: 'div', attributes: { class: 'card-title' }, children: [
								this.elementModifier.createElement({
									element: 'h2', attributes: { class: 'title' }, text: 'Edit Layout Button'
								})
							]
						}),
						{
							element: 'div', attributes: { class: 'row' }, children: [
								this.elementModifier.cell({
									element: 'input', name: 'fontSize', value: this.element.find('.user-button').css()['font-size']
								}),
								this.elementModifier.cell({
									element: 'input', name: 'backgroundcolor', value: this.element.find('.user-button').css()['background-color']
								}),
								this.elementModifier.cell({
									element: 'input', name: 'fontFamily', value: this.element.find('.user-button').css()['font-family']
								}),
								this.elementModifier.cell({
									element: 'input', name: 'color', value: this.element.find('.user-button').css()['color']
								})
							]
						}
					]
				});

				this.paneContent.makeElement({
					element: 'div', attributes: { class: 'layout-recommended-row-pane card' }, children: [
						this.elementModifier.createElement({
							element: 'div', attributes: { class: 'card-title' }, children: [
								this.elementModifier.createElement({
									element: 'h2', attributes: { class: 'title' }, text: 'Edit Layout Recommended'
								})
							]
						}),
						{
							element: 'div', attributes: { class: 'row' }, children: [
								this.elementModifier.cell({
									element: 'input', name: 'fontSize', value: this.element.find('.user-recommended').css()['font-size']
								}),
								this.elementModifier.cell({
									element: 'input', name: 'fontFamily', value: this.element.find('.user-recommended').css()['font-family']
								}),
								this.elementModifier.cell({
									element: 'input', name: 'color', value: this.element.find('.user-recommended').css()['color']
								}),
								this.elementModifier.cell({
									element: 'select', name: 'toggle', options: ['show', 'hide']
								})
							]
						}
					]
				});

				this.paneContent.makeElement({
					element: 'div', attributes: { class: 'layout-info-row-pane card' }, children: [
						this.elementModifier.createElement({
							element: 'div', attributes: { class: 'card-title' }, children: [
								this.elementModifier.createElement({
									element: 'h2', attributes: { class: 'title' }, text: 'Edit Layout Info'
								})
							]
						}),
						{
							element: 'div', attributes: { class: 'row' }, children: [
								this.elementModifier.cell({
									element: 'input', name: 'fontSize', value: this.element.find('.power-text').css()['font-size']
								}),
								this.elementModifier.cell({
									element: 'input', name: 'fontFamily', value: this.element.find('.power-text').css()['font-family']
								}),
								this.elementModifier.cell({
									element: 'input', name: 'color', value: this.element.find('.power-text').css()['color']
								}),
								this.elementModifier.cell({
									element: 'select', name: 'toggle', options: ['show', 'hide']
								})
							]
						}
					]
				});

				this.paneContent.makeElement({
					element: 'div', attributes: { class: 'layout-header-row-pane card' }, children: [
						this.elementModifier.createElement({
							element: 'div', attributes: { class: 'card-title' }, children: [
								this.elementModifier.createElement({
									element: 'h2', attributes: { class: 'title' }, text: 'Edit Layout Header'
								})
							]
						}),
						{
							element: 'div', attributes: { class: 'row' }, children: [
								this.elementModifier.cell({
									element: 'input', name: 'fontSize', value: this.element.find('.user-header').css()['font-size']
								}),
								this.elementModifier.cell({
									element: 'input', name: 'fontFamily', value: this.element.find('.user-header').css()['font-family']
								}),
								this.elementModifier.cell({
									element: 'input', name: 'color', value: this.element.find('.user-header').css()['color']
								}),
								this.elementModifier.cell({
									element: 'select', name: 'toggle', options: ['show', 'hide']
								})
							]
						}
					]
				});
				this.paneContent.append(this.generatePaneContent({ list: userListRows }));
			}
		}
		return this.paneContent;
	}

	public generatePaneContent(params) {
		let userListPane = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'card list-pane' }, children: [
				this.elementModifier.createElement({
					element: 'div', attributes: { class: 'card-title' }, children: [
						this.elementModifier.createElement({
							element: 'h2', attributes: { class: 'title' }, text: 'User'
						})
					]
				}),
			]
		});

		for (let i = 0; i < params.list.length; i++) {
			userListPane.makeElement({
				element: 'div',
				attributes: { class: 'crater-power-user-pane row' },
				children: [
					this.elementModifier.cell({
						element: 'img', name: 'image', attributes: {}, dataAttributes: { class: 'crater-icon', src: params.list[i].find('.crater-power-image').src }
					}),
					this.elementModifier.cell({
						element: 'input', name: 'header', attribute: { class: 'crater-user-header' }, value: params.list[i].find('.user-header').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'info-text', value: params.list.title || params.list[i].find('.power-text').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'recommended', value: params.list.subTitle || params.list[i].find('.user-recommended').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'button', value: params.list.body || params.list[i].find('.user-button').textContent
					})
				]
			});
		}

		return userListPane;
	}

	public listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		let self = this;
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;
		this.paneContent = this.sharePoint.app.find('.crater-property-content').monitor();

		window.onerror = (msg, url, lineNumber, columnNumber, error) => {
			console.log(msg, url, lineNumber, columnNumber, error);
		};

		if (this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.loginType.length !== 0) {

			let powerPane = this.paneContent.find('.power-pane');
			let sizePane = this.paneContent.find('.size-pane');
			if (this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.loginType.toLowerCase() === 'user') {
				powerPane.find('#user-rights-cell').parentElement.remove();
			} else {
				powerPane.find('#user-rights-cell').onChanged(value => {
					switch (value.toLowerCase()) {
						case 'allow':
							this.getAccessToken().then(response => {
								this.addUser({ accessToken: response, groupID: this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.groupId });
							});
							break;
						case 'deny':
							this.getAccessToken().then(response => {
								this.deleteUser({ accessToken: response, groupID: this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.groupId, username: this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.username });
							});
							break;
					}
				});
			}

			let changedPage = () => {
				try {
					if (powerPane.find('#page-cell')) {
						powerPane.find('#page-cell').onChanged(value => {
							for (let page = 0; page < self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.pageName.length; page++) {
								for (let property in self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.pageName[page]) {
									if (value === self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.pages[page].displayName) {
										self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.namePage = self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.pages[page].pageName;
									}
								}
							}
						});
					}
				} catch (error) {
					console.log(error.message);
				}

			};
			let changedTile = () => {
				try {
					if (powerPane.find('#tile-cell')) {
						powerPane.find('#tile-cell').onChanged(value => {

							for (let tile = 0; tile < self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.tiles.length; tile++) {
								for (let property in self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.tiles[tile]) {
									if (value === self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.tiles[tile].tileName) {
										if (powerPane.find('#filter-panel-cell')) powerPane.find('#filter-panel-cell').parentElement.parentElement.remove();
										if (powerPane.find('#navigation-cell')) powerPane.find('#navigation-cell').parentElement.parentElement.remove();
										if (powerPane.find('#page-cell')) powerPane.find('#page-cell').parentElement.remove();

										self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.tokenEmbed = 'tile';
										self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.embedType = 'tile';
										self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.tileId = self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.tiles[tile].tileId;
										self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.reportId = self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.tiles[tile].tileId;
										self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.embedUrl = self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.tiles[tile].embedUrl;
										self.getPages({ accessToken: self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.accessToken, groupID: self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.groupId, reportID: self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.tiles[tile].reportId });
										changedPage();
									}
									else if (value.toLowerCase() === 'show full dashboard') {
										self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.tokenEmbed = 'fulldashboard';
										self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.embedType = 'dashboard';
										self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.reportId = self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.dashboardId;
										self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.embedUrl = self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.dashboardEmbedUrl;
										self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.namePage = '';
										if (powerPane.find('#page-cell')) powerPane.find('#page-cell').parentElement.remove();
										if (powerPane.find('#filter-panel-cell')) powerPane.find('#filter-panel-cell').parentElement.parentElement.remove();
										if (powerPane.find('#navigation-cell')) powerPane.find('#navigation-cell').parentElement.parentElement.remove();
									}
								}
							}

							if (value.toLowerCase() !== 'show full dashboard') {
								powerPane.find('.power-embed').makeElement({
									element: 'div', children: [
										self.elementModifier.cell({
											element: 'select', name: 'page', options: self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.pageName
										})
									]
								});

								powerPane.makeElement({
									element: 'div', attributes: { class: 'row' }, children: [
										self.elementModifier.cell({
											element: 'select', name: 'filter-panel', options: ['show', 'hide']
										}),
										self.elementModifier.cell({
											element: 'select', name: 'navigation', options: ['show', 'hide']
										})
									]
								});

								powerPane.find('#filter-panel-cell').onChanged(val => {
									switch (val.toLowerCase()) {
										case 'show':
											self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.showFilter = true;
											break;
										case 'hide':
											self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.showFilter = false;
											break;
										default:
											self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.showFilter = false;
									}
								});

								powerPane.find('#navigation-cell').onChanged(val => {
									switch (val.toLowerCase()) {
										case 'show':
											self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.showNavContent = true;
											break;
										case 'hide':
											self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.showNavContent = false;
											break;
										default:
											self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.showNavContent = false;
									}
								});
							}
						});
					}
				} catch (error) {
					console.log(error.message);
				}

			};
			let changedReport = () => {
				try {
					if (powerPane.find('#view-cell')) {
						powerPane.find('#view-cell').onChanged(value => {
							if (powerPane.find('#page-cell')) powerPane.find('#page-cell').parentElement.remove();
							if (powerPane.find('#filter-panel-cell')) powerPane.find('#filter-panel-cell').parentElement.parentElement.remove();
							if (powerPane.find('#navigation-cell')) powerPane.find('#navigation-cell').parentElement.parentElement.remove();

							for (let view = 0; view < self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.reports.length; view++) {
								for (let property in self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.reports[view]) {
									if (value === self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.reports[view].reportName) {
										self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.tokenEmbed = 'reportName';
										self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.reportId = self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.reports[view].reportId;
										self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.embedUrl = self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.reports[view].embedUrl;
										self.getPages({ accessToken: self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.accessToken, groupID: self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.groupId, reportID: self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.reportId }).catch(error => console.log(error.message));
									}
								}
							}
							powerPane.find('.power-embed').makeElement({
								element: 'div', children: [
									self.elementModifier.cell({
										element: 'select', name: 'page', options: self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.pageName
									})
								]
							});
							changedPage();
							powerPane.makeElement({
								element: 'div', attributes: { class: 'row' }, children: [
									self.elementModifier.cell({
										element: 'select', name: 'filter-panel', options: ['show', 'hide']
									}),
									self.elementModifier.cell({
										element: 'select', name: 'navigation', options: ['show', 'hide']
									})
								]
							});

							powerPane.find('#filter-panel-cell').onChanged(val => {
								switch (val.toLowerCase()) {
									case 'show':
										self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.showFilter = true;

										break;
									case 'hide':
										self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.showFilter = false;
										break;
									default:
										self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.showFilter = false;
								}
							});

							powerPane.find('#navigation-cell').onChanged(val => {
								switch (val.toLowerCase()) {
									case 'show':
										self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.showNavContent = true;
										break;
									case 'hide':
										self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.showNavContent = false;
										break;
									default:
										self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.showNavContent = false;
								}
							});
						});
					}
				} catch (error) {
					console.log(error.message);
				}

			};
			let changedDashboard = () => {
				try {
					if (powerPane.find('#view-cell')) {
						powerPane.find('#view-cell').onChanged(value => {
							if (powerPane.find('#filter-panel-cell')) powerPane.find('#filter-panel-cell').parentElement.parentElement.remove();
							if (powerPane.find('#navigation-cell')) powerPane.find('#navigation-cell').parentElement.parentElement.remove();
							if (powerPane.find('#page-cell')) powerPane.find('#page-cell').parentElement.remove();
							if (powerPane.find('#tile-cell')) powerPane.find('#tile-cell').parentElement.remove();
							for (let view = 0; view < self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.dashboards.length; view++) {
								for (let property in self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.dashboards[view]) {
									if (value === self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.dashboards[view].dashboardName) {
										self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.tokenEmbed = 'fulldashboard';
										self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.dashboardId = self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.dashboards[view].dashboardId;
										self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.dashboardEmbedUrl = self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.dashboards[view].embedUrl;
										self.getTiles(self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.accessToken, self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.groupId, self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.dashboardId);
									}
								}
							}
							powerPane.find('.power-embed').makeElement({
								element: 'div', children: [
									self.elementModifier.cell({
										element: 'select', name: 'tile', options: self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.tileName
									})
								]
							});
							self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.embedType = 'dashboard';
							changedTile();
						});
					}
				} catch (error) {
					console.log(error.message);
				}

			};

			powerPane.find('#WorkSpace-cell').onChanged(value => {
				this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.changed = true;
				if (powerPane.find('#embed-type-cell')) powerPane.find('#embed-type-cell').parentElement.remove();
				if (powerPane.find('#view-cell')) powerPane.find('#view-cell').parentElement.remove();
				if (powerPane.find('#tile-cell')) powerPane.find('#tile-cell').parentElement.remove();
				if (powerPane.find('#page-cell')) powerPane.find('#page-cell').parentElement.remove();
				if (powerPane.find('#filter-panel-cell')) powerPane.find('#filter-panel-cell').parentElement.parentElement.remove();
				if (powerPane.find('#navigation-cell')) powerPane.find('#navigation-cell').parentElement.parentElement.remove();
				for (let group = 0; group < this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.groups.length; group++) {
					for (let property in this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.groups[group]) {
						if (value === this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.groups[group].workspaceName) {
							this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.groupId = this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.groups[group].workspaceId;
							this.getReports(this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.accessToken, this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.groupId);
							this.getDashboards(this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.accessToken, this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.groupId);
						} else if (value.toLowerCase() === 'my workspace') {
							this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.groupId = 'none';
							this.getReports(this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.accessToken, this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.groupId, 'third');
							this.getDashboards(this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.accessToken, this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.groupId, 'third');
						}
					}
				}

				powerPane.find('.power-embed').makeElement({
					element: 'div', children: [
						self.elementModifier.cell({
							element: 'select', name: 'embed-type', options: ['Dashboard', 'Report']
						})
					]
				});

				powerPane.find('#embed-type-cell').onChanged(val => {
					switch (val.toLowerCase()) {
						case 'report':
							if (powerPane.find('#filter-panel-cell')) powerPane.find('#filter-panel-cell').parentElement.parentElement.remove();
							if (powerPane.find('#navigation-cell')) powerPane.find('#navigation-cell').parentElement.parentElement.remove();
							if (powerPane.find('#page-cell')) powerPane.find('#page-cell').parentElement.remove();
							if (powerPane.find('#view-cell')) powerPane.find('#view-cell').parentElement.remove();
							if (powerPane.find('#tile-cell')) powerPane.find('#tile-cell').parentElement.remove();
							powerPane.find('.power-embed').makeElement({
								element: 'div', children: [
									this.elementModifier.cell({
										element: 'select', name: 'view', options: this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.reportName
									})
								]
							});
							self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.tileId = '';
							self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.dashboardId = '';
							// if (this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.reportName.indexOf('No Reports') !== -1) {
							// 	powerPane.find('#view-cell').options[0].selected = true;
							// 	powerPane.find('#view-cell').disabled = true;
							// }
							this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.embedType = val.toLowerCase();
							changedReport();
							break;
						case 'dashboard':
							if (powerPane.find('#filter-panel-cell')) powerPane.find('#filter-panel-cell').parentElement.parentElement.remove();
							if (powerPane.find('#navigation-cell')) powerPane.find('#navigation-cell').parentElement.parentElement.remove();
							if (powerPane.find('#page-cell')) powerPane.find('#page-cell').parentElement.remove();
							if (powerPane.find('#view-cell')) powerPane.find('#view-cell').parentElement.remove();
							if (powerPane.find('#tile-cell')) powerPane.find('#tile-cell').parentElement.remove();

							powerPane.find('.power-embed').makeElement({
								element: 'div', children: [
									this.elementModifier.cell({
										element: 'select', name: 'view', options: this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.dashBoardName
									})
								]
							});
							// if (this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.dashBoardName.indexOf('No Dashboards') !== -1) {
							// 	powerPane.find('#view-cell').options[0].selected = true;
							// 	powerPane.find('#view-cell').disabled = true;
							// }
							this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.embedType = val.toLowerCase();
							changedDashboard();
							break;
					}
				});
			});

			sizePane.find('#display-cell').onChanged(value => {
				switch (value) {
					case '16:9 (1280px x 720px)':
						if (sizePane.find('#width-cell')) sizePane.find('#width-cell').parentElement.remove();
						if (sizePane.find('#height-cell')) sizePane.find('#height-cell').parentElement.remove();
						this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.width = '1280px';
						this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.height = '720px';
						break;
					case '4:3 (1000px x 750px)':
						if (sizePane.find('#width-cell')) sizePane.find('#width-cell').parentElement.remove();
						if (sizePane.find('#height-cell')) sizePane.find('#height-cell').parentElement.remove();
						this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.width = '1000px';
						this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.height = '750px';
						break;
					case 'Custom Size':
						sizePane.find('.power-embed').makeElement({
							element: 'div', children: [
								this.elementModifier.cell({
									element: 'input', name: 'width', value: this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.width
								}),
								this.elementModifier.cell({
									element: 'input', name: 'height', value: this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.height
								})
							]
						});
						sizePane.find('#width-cell').onChanged(val => {
							this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.width = val;
						});
						sizePane.find('#height-cell').onChanged(val => {
							this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.height = val;
						});
						break;
				}
			});
		}
		else {
			let layoutRowPane = this.paneContent.find('.layout-pane');
			let layoutButtonRowPane = this.paneContent.find('.layout-button-row-pane');
			let layoutRecommendedRowPane = this.paneContent.find('.layout-recommended-row-pane');
			let layoutInfoRowPane = this.paneContent.find('.layout-info-row-pane');
			let layoutHeaderRowPane = this.paneContent.find('.layout-header-row-pane');
			let userList = this.element.find('.crater-power-connect');
			let userListRow = userList.findAll('.user');

			let layoutBackgroundParent = layoutRowPane.find('#backgroundcolor-cell').parentNode;
			this.pickColor({ parent: layoutBackgroundParent, cell: layoutBackgroundParent.find('#backgroundcolor-cell') }, (backgroundColor) => {
				this.element.findAll('.login-container').forEach(element => {
					element.css({ backgroundColor });
				});
				layoutBackgroundParent.find('#backgroundcolor-cell').value = backgroundColor;
				layoutBackgroundParent.find('#backgroundcolor-cell').setAttribute('value', backgroundColor); //set the value of the eventColor cell in the pane to the color
			});

			let layoutButtonParent = layoutButtonRowPane.find('#backgroundcolor-cell').parentNode;
			this.pickColor({ parent: layoutButtonParent, cell: layoutButtonParent.find('#backgroundcolor-cell') }, (backgroundColor) => {
				this.element.findAll('.user-button').forEach(element => {
					element.css({ backgroundColor });
				});
				layoutButtonParent.find('#backgroundcolor-cell').value = backgroundColor;
				layoutButtonParent.find('#backgroundcolor-cell').setAttribute('value', backgroundColor); //set the value of the eventColor cell in the pane to the color
			});

			let layoutButtonColor = layoutButtonRowPane.find('#color-cell').parentNode;
			this.pickColor({ parent: layoutButtonColor, cell: layoutButtonColor.find('#color-cell') }, (color) => {
				this.element.findAll('.user-button').forEach(element => {
					element.css({ color });
				});
				layoutButtonColor.find('#color-cell').value = color;
				layoutButtonColor.find('#color-cell').setAttribute('value', color);
			});

			let layoutRecommendedColor = layoutRecommendedRowPane.find('#color-cell').parentNode;
			this.pickColor({ parent: layoutRecommendedColor, cell: layoutRecommendedColor.find('#color-cell') }, (color) => {
				this.element.findAll('.user-recommended').forEach(element => {
					element.css({ color });
				});
				layoutRecommendedColor.find('#color-cell').value = color;
				layoutRecommendedColor.find('#color-cell').setAttribute('value', color);
			});

			let layoutInfoColor = layoutInfoRowPane.find('#color-cell').parentNode;
			this.pickColor({ parent: layoutInfoColor, cell: layoutInfoColor.find('#color-cell') }, (color) => {
				this.element.findAll('.power-text').forEach(element => {
					element.css({ color });
				});
				layoutInfoColor.find('#color-cell').value = color;
				layoutInfoColor.find('#color-cell').setAttribute('value', color);
			});

			let layoutHeaderColor = layoutHeaderRowPane.find('#color-cell').parentNode;
			this.pickColor({ parent: layoutHeaderColor, cell: layoutHeaderColor.find('#color-cell') }, (color) => {
				this.element.findAll('.user-header').forEach(element => {
					element.css({ color });
				});
				layoutHeaderColor.find('#color-cell').value = color;
				layoutHeaderColor.find('#color-cell').setAttribute('value', color);
			});

			layoutButtonRowPane.find('#fontSize-cell').onChanged(value => {
				this.element.findAll('.user-button').forEach(element => {
					element.css({ fontSize: value });
				});
			});

			layoutRecommendedRowPane.find('#fontSize-cell').onChanged(value => {
				this.element.findAll('.user-recommended').forEach(element => {
					element.css({ fontSize: value });
				});
			});

			layoutInfoRowPane.find('#fontSize-cell').onChanged(value => {
				this.element.findAll('.power-text').forEach(element => {
					element.css({ fontSize: value });
				});
			});

			layoutHeaderRowPane.find('#fontSize-cell').onChanged(value => {
				this.element.findAll('.user-header').forEach(element => {
					element.css({ fontSize: value });
				});
			});

			layoutButtonRowPane.find('#fontFamily-cell').onChanged(value => {
				this.element.findAll('.user-button').forEach(element => {
					element.css({ fontFamily: value });
				});
			});

			layoutRecommendedRowPane.find('#fontFamily-cell').onChanged(value => {
				this.element.findAll('.user-recommended').forEach(element => {
					element.css({ fontFamily: value });
				});
			});

			layoutInfoRowPane.find('#fontFamily-cell').onChanged(value => {
				this.element.findAll('.power-text').forEach(element => {
					element.css({ fontFamily: value });
				});
			});

			layoutHeaderRowPane.find('#fontFamily-cell').onChanged(value => {
				this.element.findAll('.user-header').forEach(element => {
					element.css({ fontFamily: value });
				});
			});

			let showRecommended = layoutRecommendedRowPane.find('#toggle-cell');
			showRecommended.addEventListener('change', e => {

				switch (showRecommended.value.toLowerCase()) {
					case "hide":
						this.element.findAll('.user-recommended').forEach(element => {
							element.style.display = "none";
						});
						break;
					case "show":
						this.element.findAll('.user-recommended').forEach(element => {
							element.style.display = "block";
						});
						break;
					default:
						this.element.findAll('.user-recommended').forEach(element => {
							element.style.display = "none";
						});
				}
			});

			let showInfo = layoutInfoRowPane.find('#toggle-cell');
			showInfo.addEventListener('change', e => {

				switch (showInfo.value.toLowerCase()) {
					case "hide":
						this.element.findAll('.power-text').forEach(element => {
							element.style.display = "none";
						});
						break;
					case "show":
						this.element.findAll('.power-text').forEach(element => {
							element.style.display = "block";
						});
						break;
					default:
						this.element.findAll('.power-text').forEach(element => {
							element.style.display = "none";
						});
				}
			});

			let showHeader = layoutHeaderRowPane.find('#toggle-cell');
			showHeader.addEventListener('change', e => {

				switch (showHeader.value.toLowerCase()) {
					case "hide":
						this.element.findAll('.user-header').forEach(element => {
							element.style.display = "none";
						});
						break;
					case "show":
						this.element.findAll('.user-header').forEach(element => {
							element.style.display = "block";
						});
						break;
					default:
						this.element.findAll('.user-header').forEach(element => {
							element.style.display = "none";
						});
				}
			});

			let userRowHandler = (userRowPane, userRowDom) => {
				let iconParent = userRowPane.find('#image-cell').parentNode;
				this.uploadImage({ parent: iconParent }, (image) => {
					iconParent.find('#image-cell').src = image.src;
					this.element.find('.crater-power-image').src = image.src;
				});
				userRowPane.find('#header-cell').onChanged(value => {
					userRowDom.find('.user-header').innerHTML = value;
				});

				userRowPane.find('#info-text-cell').onChanged(value => {
					userRowDom.find('.power-text').innerHTML = value;
				});

				userRowPane.find('#recommended-cell').onChanged(value => {
					userRowDom.find('.user-recommended').innerHTML = value;
				});

				userRowPane.find('#button-cell').onChanged(value => {
					userRowDom.find('.user-button').innerHTML = value;
				});

			};

			let paneItems = this.paneContent.findAll('.crater-power-user-pane');
			paneItems.forEach((userRow, position) => {
				userRowHandler(userRow, userListRow[position]);
			});

		}

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
			let draftPower = this.sharePoint.properties.pane.content[this.key].settings.myPowerBi;

			if (this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.changed) {
				if (this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.groupId !== 'none') {
					this.getEmbedToken({ accessToken: draftPower.accessToken, groupID: draftPower.groupId, reportID: draftPower.reportId, generateUrl: draftPower.tokenEmbed }).then(response => {
						this.element.find('.crater-power-container').css({ width: draftPower.width, height: draftPower.height });
						this.element.find('#renderContainer').css({ width: '100%', height: draftPower.height });
						this.embedPower({ accessToken: response, type: draftPower.embedType, embedUrl: draftPower.embedUrl });
					});
				} else {
					if (draftPower.embedType === 'report') {
						if (this.element.find('#renderContainer')) this.element.find('#renderContainer').remove();
						let powerContainer = this.element.find('.crater-power-container') as any;
						powerContainer.makeElement({
							element: 'div', attributes: { id: 'renderContainer' }, children: [
								{
									element: 'div', attributes: { id: 'render-error' }, children: [
										{ element: 'p', attributes: { id: 'render-text' } }
									]
								}
							]
						});
						let render = powerContainer.find('#renderContainer');
						const iframeURL = `${draftPower.embedUrl}&autoAuth=true&ctid=${draftPower.tenantID}&filterPaneEnabled=${draftPower.showFilter}&navContentPaneEnabled=${draftPower.showNavContent}&pageName=${draftPower.namePage}`;
						render.makeElement({
							element: 'div', attributes: { class: 'power-iframe' }, children: [
								{
									element: 'iframe', attributes: { width: draftPower.width, height: draftPower.height, src: iframeURL, frameborder: "0", allowFullScreen: "true" }
								}
							]
						});
						this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.embedded = false;

						render.find('#render-error').style.display = "none";
						render.style.display = "block";
					}
				}
			} else {
				if (this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.groupId !== 'none') {
					this.element.find('.crater-power-container').css({ width: draftPower.width, height: draftPower.height });
					this.element.find('#renderContainer').css({ width: '100%', height: draftPower.height });
				} else {
					let powerIframe = this.element.find('#renderContainer').find('iframe');
					powerIframe.width = draftPower.width;
					powerIframe.height = draftPower.height;
				}

			}
		});
	}
}

class Birthday extends CraterWebParts {
	public key: any;
	public element: any;
	public paneContent: any;
	public elementModifier: any = new ElementModifier();
	public params: any;

	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.params = params;
	}

	public render(params) {
		let birthday = this.createKeyedElement({
			element: 'div', attributes: { class: 'crater-birthday crater-component', 'data-type': 'birthday' }, children: [
				{
					element: 'div', attributes: { class: 'crater-birthday-header' }, children: [
						{ element: 'img', attributes: { class: 'crater-birthday-header-img', src: this.sharePoint.images.append } },
						{ element: 'div', attributes: { class: 'crater-birthday-header-title' }, text: 'Birthdays Coming Up...' }
					]
				},
				{
					element: 'div', attributes: { class: 'crater-birthday-div' }
				},
				{
					element: 'div', attributes: { class: 'birthday-loading' }, children: [

						{ element: 'p', attributes: { class: 'birthday-loading-text' }, text: 'Fetching Employee Directory, Please wait...' },
						{
							element: 'img', attributes: { class: 'birthday-loading-image crater-icon', src: this.sharePoint.images.loading, style: { width: '20px', height: '20px' } }
						}
					]
				}
			]
		});

		this.key = this.key || birthday.dataset.key;
		this.sharePoint.properties.pane.content[this.key].settings.myBirthdays = {
			fetched: false,
			mode: 'Birthday',
			interval: 999,
			count: 0,
			sortBy: 'Birthday'
		};

		return birthday;
	}

	public rendered(params) {
		this.key = params.element.dataset.key;
		this.element = params.element;
		this.getUser();

		params.element.find('.birthday-loading-text').textContent = 'Fetching Employee Directory, Please wait...';

		if (!this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.fetched) {
			setTimeout(() => {
				if (this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.users.length !== 0) {
					let renderedCount = (this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.count) ? this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.count : this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.users.length;
					let birthdayDiv = params.element.find('.crater-birthday-div');
					let birthdayLoading = params.element.find('.birthday-loading') as any;
					birthdayLoading.style.display = 'none';
					if (this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.mode.toLowerCase() === 'birthday') {
						for (let i = 0; i < renderedCount; i++) {
							const phoneExists = (this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.users[i].phone) ? `<a href="tel:${this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.users[i].phone}">
							<img class="img" src="http://www.free-icons-download.net/images/call-icon-16876.png"
								alt=""></a>
								<a href="sms:${this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.users[i].phone}">
							<img class="img" src="http://www.free-icons-download.net/images/bubble-message-icon-74535.png"
								alt=""></a>
								` : '';

							let birthdayHTML = `<div class="birthday" divid="${this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.users[i].id}">
										<div class="birthday-image">
											<img class="name-image1" 
												src=${this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.users[i].image}
												alt="">
										</div>
										<div class="div">
											<div class="name"><span class="personName">${this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.users[i].displayName}</span>
												<img class="name-image" src="https://img.icons8.com/material-sharp/24/000000/birthday.png"
													alt=""><br>
											</div>
											<span class="title">${this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.users[i].mail}</span>
											<div class="date"><span class="birthDate">${new Date(this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.users[i].birthDate).toString().split(`${new Date(this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.users[i].birthDate).getFullYear()}`)[0]}</span>
												<a href="mailto:${this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.users[i].mail}?Subject=Happy%20Birthday%20${this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.users[i].firstName}">
													<img class="img" src="https://img.icons8.com/material-two-tone/24/000000/composing-mail.png"
														alt=""></a>
													${phoneExists}	
											</div>
										</div>
								</div>`;

							birthdayDiv.innerHTML += birthdayHTML;
						}
					} else if (this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.mode.toLowerCase() === 'anniversary') {
						for (let i = 0; i < renderedCount; i++) {
							const phoneExists = (this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.users[i].phone) ? `<a href="tel:${this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.users[i].phone}">
							<img class="img" src="http://www.free-icons-download.net/images/call-icon-16876.png"
								alt=""></a>
								<a href="sms:${this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.users[i].phone}">
							<img class="img" src="http://www.free-icons-download.net/images/bubble-message-icon-74535.png"
								alt=""></a>
								` : '';
							let birthdayHTML = `<div class="birthday" divid="${this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.users[i].id}">
									<div class="birthday-image">
										<img class="name-image1"
											src=${this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.users[i].image}
											alt="">
									</div>
									<div class="div">
										<div class="name"><span class="personName">${this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.users[i].displayName}</span>
											<img class="name-image" src="https://img.icons8.com/material-sharp/24/000000/birthday.png"
												alt=""><br>
										</div>
										<span class="title">${this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.users[i].mail}</span>
										<div class="date"><span class="birthDate">${this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.users[i].anniversary}</span>
											<a href="mailto:${this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.users[i].mail}?Subject=Happy%20Birthday%20${this.sharePoint.properties.pane.content[this.key].settings.myBirthdays.users[i].firstName}">
												<img class="img" src="https://img.icons8.com/material-two-tone/24/000000/composing-mail.png"
													alt=""></a>
											${phoneExists}
										</div>
									</div>
								</div>`;
							birthdayDiv.innerHTML += birthdayHTML;
						}
					}
				} else {
					params.element.find('.birthday-loading-text').textContent = 'Taking longer than expected. Please check your internet connection';
				}
			}, 12000);
		}

		window.onerror = (msg, url, lineNumber, columnNumber, error) => {
			console.log(msg, url, lineNumber, columnNumber, error);
		};
	}

	public getDAYS(x) {
		let y = 365;
		let y2 = 31;
		let remainder = x % y;
		let casio = remainder % y2;
		let year = (x - remainder) / y;
		let month = (remainder - casio) / y2;

		let result = year + " Year(s)" + ", " + month + " Month(s)" + ", and " + casio + " Day(s)";

		return {
			result,
			year,
			month,
			days: casio
		};
	}

	public getUser = async () => {
		let draftBirthday = this.sharePoint.properties.pane.content[this.key].settings.myBirthdays;
		draftBirthday.users = [];
		let self = this;

		let getImage = (imgID, element?) => {
			const setSource = (element) ? element : this.element;
			this.sharePoint.connection.getWithGraph().then((client) => {
				client.api(`/users/${imgID}/photo/$value`)
					.responseType('blob')
					.get((error: any, result: any, rawResponse?: any) => {
						if (!func.setNotNull(result)) return;
						if (draftBirthday.users.length !== 0) {
							for (let p = 0; p < draftBirthday.users.length; p++) {
								if (draftBirthday.users[p].id === imgID) {
									const myBlob = new Blob([result], { type: 'blob' });
									const blobUrl = URL.createObjectURL(myBlob);
									draftBirthday.users[p].photo = result;
									draftBirthday.users[p].image = blobUrl;
									if ((setSource.find('.birthday')) && (!setSource.find('.no-users'))) {
										let renderedBirthdays = setSource.findAll('.birthday') as any;
										if (renderedBirthdays[p].getAttribute('divid') === imgID) {
											renderedBirthdays[p].find('.name-image1').src = blobUrl;
										}
									}
								}
							}
						}
					});
			});
		};

		let getDepartment = (dptID) => {
			this.sharePoint.connection.getWithGraph().then(client => {
				client.api(`/users/${dptID}/department`)
					.get((error: any, result: any, rawResponse?: any) => {
						if (!func.setNotNull(result)) return;
						if (draftBirthday.users.length !== 0) {
							for (let p = 0; p < draftBirthday.users.length; p++) {
								if (draftBirthday.users[p].id === dptID) {
									draftBirthday.users[p].department = result['value'];
								}
							}
						}
					});
			});
		};

		let getBirthday = (bthID) => {
			this.sharePoint.connection.getWithGraph().then((client) => {
				client.api(`/users/${bthID}/birthday`)
					.get((error: any, result: any, rawResponse?: any) => {
						if (!func.setNotNull(result)) return;
						if (draftBirthday.users.length !== 0) {
							for (let p = 0; p < draftBirthday.users.length; p++) {
								if (draftBirthday.users[p].id === bthID) {
									const birthday = result['value'];
									if (birthday.indexOf('T') != -1) {
										let year = birthday.split('T')[0].split('-')[0];
										let month = birthday.split('T')[0].split('-')[1];
										let day = birthday.split('T')[0].split('-')[2];
										let newDate = month + '/' + day + '/' + year;

										draftBirthday.users[p].birthDate = newDate;
										if (draftBirthday.mode.toLowerCase() === 'birthday') {
											if ((this.element.find('.birthday')) && (!this.element.find('.no-users'))) {
												let renderedBirthdays = this.element.findAll('.birthday') as any;
												if (renderedBirthdays[p].getAttribute('divid') === bthID) {
													renderedBirthdays[p].find('.birthDate').textContent = new Date(draftBirthday.users[p].birthDate).toString().split(`${new Date(draftBirthday.users[p].birthDate).getFullYear()}`)[0];
												}
											}
										}
									}
								}
							}
						}
					});
			});
		};

		let getHireDate = (hID) => {
			this.sharePoint.connection.getWithGraph().then(client => {
				client.api(`/users/${hID}/hireDate`)
					.get((error: any, result: any, rawResponse?: any) => {
						if (!func.setNotNull(result)) return;
						if (draftBirthday.users.length !== 0) {
							for (let p = 0; p < draftBirthday.users.length; p++) {
								if (draftBirthday.users[p].id === hID) {
									const hired = result['value'];
									if (hired.indexOf('T') != -1) {
										let hYear = hired.split('T')[0].split('-')[0];
										let hMonth = hired.split('T')[0].split('-')[1];
										let hDay = hired.split('T')[0].split('-')[2];
										let hDate = hMonth + '/' + hDay + '/' + hYear;

										draftBirthday.users[p].hireDate = hDate;

										let date = new Date();
										//@ts-ignore
										let personHireDate = new Date(draftBirthday.users[p].hireDate);
										let timeLeft = date.getTime() - personHireDate.getTime();

										let daysLeft = Math.floor(timeLeft / (1000 * 60 * 60 * 24));
										let hoursLeft = Math.floor((timeLeft % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
										let minutesLeft = Math.floor((timeLeft % (1000 * 60 * 60)) / (1000 * 60));
										let secondsLeft = Math.floor((timeLeft % (1000 * 60)) / (1000));

										draftBirthday.users[p].anniversary = self.getDAYS(daysLeft).result;
										draftBirthday.users[p].year = self.getDAYS(daysLeft).year;

										if (draftBirthday.mode.toLowerCase() === 'anniversary') {
											if ((this.element.find('.birthday')) && (!this.element.find('.no-users'))) {
												let renderedBirthdays = this.element.findAll('.birthday') as any;
												if (renderedBirthdays[p].getAttribute('divid') === hID) {
													renderedBirthdays[p].find('.birthDate').textContent = draftBirthday.users[p].anniversary;
												}
											}
										}

									}
								}
							}
						}
					});
			});
		};

		this.sharePoint.connection.getWithGraph().then(client => {
			client.api('/users')
				.select('mail, displayName, givenName, id, surname, jobTitle, mobilePhone, officeLocation, photo, image')
				.get(async (error: any, result: MicrosoftGraph.User, rawResponse?: any) => {
					for (let p = 0; p < result['value'].length; p++) {
						draftBirthday.users.push({
							id: result['value'][p].id,
							displayName: result['value'][p].displayName,
							firstName: result['value'][p].givenName,
							lastName: result['value'][p].surname,
							mail: result['value'][p].mail,
							Title: result['value'][p].jobTitle,
							phone: result['value'][p].mobilePhone,
							birthDate: '01/01/1985',
							hireDate: '01/01/0001',
							image: 'http://icons.iconarchive.com/icons/graphicloads/flat-finance/72/person-icon.png'
						});

						if (draftBirthday.users.length != 0) {
							draftBirthday.fetched = true;
							getImage(result['value'][p].id);
							getDepartment(result['value'][p].id);
							getBirthday(result['value'][p].id);
							getHireDate(result['value'][p].id);
						}
					}
				});
		});
	}

	public setUpPaneContent(params) {
		let key = params.element.dataset['key'];//create a key variable and set it to the webpart key
		this.element = this.sharePoint.properties.pane.content[key].draft.dom;//define the declared element to the draft dom content
		let draftBirthday = this.sharePoint.properties.pane.content[key].settings.myBirthdays;

		this.paneContent = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'crater-property-content' }
		}).monitor(); //monitor the content pane 
		if (this.sharePoint.properties.pane.content[key].draft.pane.content.length !== 0) {//check if draft pane content is not empty and set it to the pane content
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[key].draft.pane.content;
		}
		else if (this.sharePoint.properties.pane.content[key].content.length !== 0) {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[key].content;
		} else {
			let birthdayList = this.sharePoint.properties.pane.content[key].draft.dom.find('.crater-birthday-div');
			let birthdayListRows = birthdayList.findAll('.birthday');

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'title-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							{ element: 'h2', attributes: { class: 'title' }, text: 'Customize Birthday' }
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							{
								element: 'div', attributes: { class: 'message-note' }, children: [
									{
										element: 'div', attributes: { class: 'message-text' }, children: [
											{ element: 'p', attributes: { style: { color: 'green' } }, text: `MODE: Birthday/Anniversary.` },
											{ element: 'p', attributes: { style: { color: 'green' } }, text: `INTERVAL: In Anniversary Mode, enter “999” to show all anniversaries, or enter a number to only show employees that have the set anniversary (e.g 10 -> 10 years)` },
											{ element: 'p', attributes: { style: { color: 'green' } }, text: `COUNT: The number of items to display.` },
											{ element: 'p', attributes: { style: { color: 'green' } }, text: `DAYSFUTURE: enter the number of days into the future (starting from the current date) to include in the list.` },
											{ element: 'p', attributes: { style: { color: 'green' } }, text: `DAYSPAST: enter the number of days to keep the birthday/anniversary in the list after it has passed.` }
										]
									}
								]
							}
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'title-settings row' }, children: [
							this.elementModifier.cell({
								element: 'select', name: 'mode', options: ['Birthday', 'Anniversary'], value: this.sharePoint.properties.pane.content[key].settings.myBirthdays.mode
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Interval', value: this.sharePoint.properties.pane.content[key].settings.myBirthdays.interval
							}),
							this.elementModifier.cell({
								element: 'input', name: 'daysFuture', value: ''
							}),
							this.elementModifier.cell({
								element: 'input', name: 'daysPast', value: ''
							}),
							this.elementModifier.cell({
								element: 'input', name: 'count', value: this.sharePoint.properties.pane.content[key].settings.myBirthdays.count || ''
							}),
							this.elementModifier.cell({
								element: 'select', name: 'sortBy', options: ['Name', 'Birthday', 'Anniversary'], value: this.sharePoint.properties.pane.content[key].settings.myBirthdays.sortBy
							}),
							this.elementModifier.cell({
								element: 'select', name: 'order', options: ['Ascending', 'Descending']
							})
						]
					})
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'section-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							{ element: 'h2', attributes: { class: 'title' }, text: 'Edit Section' }
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'section-height', value: this.element.find('.crater-birthday-div').css()['max-height']
							})
						]
					})
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'header-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							{ element: 'h2', attributes: { class: 'title' }, text: 'Edit Header' }
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'Image', value: this.element.find('.crater-birthday-header-img').src
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontSize', value: this.element.find('.crater-birthday-header-title').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'select', name: 'fontFamily', options: ['Comic Sans MS', 'Impact', 'Bookman', 'Garamond', 'Palatino', 'Georgia', 'Verdana', 'Times New Roman', 'Arial'], value: this.element.find('.crater-birthday-header-title').css()['font-family']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'color', value: this.element.find('.crater-birthday-header-title').css()['color']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'background-color', value: this.element.find('.crater-birthday-header').css()['color']
							}),
							this.elementModifier.cell({
								element: 'select', name: 'toggle', options: ['Show', 'Hide']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'text', value: this.element.find('.crater-birthday-header-title').textContent
							}),
							this.elementModifier.cell({
								element: 'input', name: 'header-height', value: this.element.find('.crater-birthday-header').css()['height']
							})
						]
					})
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'name-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							{ element: 'h2', attributes: { class: 'title' }, text: 'Edit Name' }
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'fontSize', value: this.element.find('.personName').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'select', name: 'fontFamily', options: ['Comic Sans MS', 'Impact', 'Bookman', 'Garamond', 'Palatino', 'Georgia', 'Verdana', 'Times New Roman', 'Arial'], value: this.element.find('.personName').css()['font-family']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'color', value: this.element.find('.personName').css()['color']
							})
						]
					})
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'job-title-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							{ element: 'h2', attributes: { class: 'title' }, text: 'Edit Title Font' }
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'fontSize', value: this.element.find('.title').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'select', name: 'fontFamily', options: ['Comic Sans MS', 'Impact', 'Bookman', 'Garamond', 'Palatino', 'Georgia', 'Verdana', 'Times New Roman', 'Arial'], value: this.element.find('.title').css()['font-family']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'color', value: this.element.find('.title').css()['color']
							}),
						]
					})
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'birthday-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							{ element: 'h2', attributes: { class: 'title' }, text: 'Edit BirthDay Font' }
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'fontSize', value: this.element.find('.birthDate').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'select', name: 'fontFamily', options: ['Comic Sans MS', 'Impact', 'Bookman', 'Garamond', 'Palatino', 'Georgia', 'Verdana', 'Times New Roman', 'Arial'], value: this.element.find('.birthDate').css()['font-family']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'color', value: this.element.find('.birthDate').css()['color']
							})
						]
					})
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'update-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							{ element: 'h2', attributes: { class: 'title' }, text: 'Update Directory' }
						]
					}),
					this.elementModifier.createElement(
						{
							element: 'div', attributes: { class: 'row' }, children: [
								this.elementModifier.createElement({
									element: 'div', attributes: { class: 'message-note' }, children: [
										{
											element: 'div', attributes: { class: 'message-text' }, children: [
												{ element: 'p', attributes: { id: 'update-message', style: { color: 'green' } }, text: `Click this button to update the user directory` },
											]
										}
									]
								}),
								this.elementModifier.createElement(
									{ element: 'button', attributes: { id: 'fetch-birthday', class: 'user-button', style: { margin: '0 auto !important' } }, text: '' }
								)
							]
						}
					)
				]
			});
			if (draftBirthday.fetched) {
				this.paneContent.find('#update-message').textContent = 'Directory Updated! Sort the list to view it';
				this.paneContent.find('#fetch-birthday').innerHTML = 'UPDATED';
			} else {
				this.paneContent.find('#update-message').textContent = 'Click this button to update the user directory';
				this.paneContent.find('#fetch-birthday').innerHTML = 'UPDATE';
			}
			this.paneContent.find('.title-settings').find('#Interval-cell').readOnly = true;

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'user-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							{ element: 'h2', attributes: { class: 'title' }, text: 'OPTIONS' }
						]
					}),

					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'display-options row' }, children: [
							this.elementModifier.cell({
								element: 'select', name: 'nameField', options: ['id', 'FullName', 'FirstName', 'LastName', 'Mail', 'Title', 'Phone', 'Birthday/Anniversary']
							}),
							this.elementModifier.cell({
								element: 'select', name: 'titleField', options: ['id', 'FullName', 'FirstName', 'LastName', 'Mail', 'Title', 'Phone', 'Birthday/Anniversary']
							}),
							this.elementModifier.cell({
								element: 'select', name: 'birthdayField', options: ['id', 'FullName', 'FirstName', 'LastName', 'Mail', 'Title', 'Phone', 'Birthday/Anniversary']
							})
						]
					})
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'background-color-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							{ element: 'h2', attributes: { class: 'title' }, text: 'Change Row Background' }
						]
					}),

					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'message-note' }, children: [
							{
								element: 'div', attributes: { class: 'message-text' }, children: [
									{ element: 'p', attributes: { id: 'update-message', style: { color: 'green' } }, text: `Enter the position of the birthday card e.g 1` },
									{ element: 'p', attributes: { id: 'update-message', style: { color: 'green' } }, text: `If you wish to alter multiple positions enter the positions as 1, 2, 3 etc...` }
								]
							}
						]
					}),

					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'display-options row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'position'
							}),
							this.elementModifier.cell({
								element: 'input', name: 'backgroundColor'
							})
						]
					})
				]
			});

		}


		return this.paneContent;
	}

	public generatePaneContent(params) {
		let birthdayListPane = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'card user-birthday-pane' }, children: [
				this.elementModifier.createElement({
					element: 'div', attributes: { class: 'card-title' }, children: [
						this.elementModifier.createElement({
							element: 'h2', attributes: { class: 'title' }, text: 'Background Color'
						})
					]
				}),
			]
		});

		for (let i = 0; i < params.list.length; i++) {
			birthdayListPane.makeElement({
				element: 'div',
				attributes: { class: 'crater-birthday-item-pane row' },
				children: [
					this.elementModifier.cell({
						element: 'input', name: 'birthdayBackground', value: params.list[i].css()['background-color']
					}),
					this.elementModifier.cell({
						element: 'input', name: 'birthdayName', value: params.list[i].find('.personName').textContent
					})
				]
			});
		}

		return birthdayListPane;


	}

	public listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		this.paneContent = this.sharePoint.app.find('.crater-property-content').monitor();
		let draftBirthday = this.sharePoint.properties.pane.content[this.key].settings.myBirthdays;
		let birthdayList = this.sharePoint.properties.pane.content[this.key].draft.dom.find('.crater-birthday-div');
		let birthdayListRow = birthdayList.findAll('.birthday') as any;
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;
		draftBirthday.sortUsers = [];

		let titlePane = this.paneContent.find('.title-pane');
		let optionsPane = this.paneContent.find('.user-pane');
		let namePane = this.paneContent.find('.name-pane');
		let jobPane = this.paneContent.find('.job-title-pane');
		let birthdayPane = this.paneContent.find('.birthday-pane');
		let headerPane = this.paneContent.find('.header-pane');
		let sectionPane = this.paneContent.find('.section-pane');
		let bgColorPane = this.paneContent.find('.background-color-pane');

		let orderCell = titlePane.find('#order-cell');
		let sortCell = titlePane.find('#sortBy-cell');
		let modeCell = titlePane.find('#mode-cell');
		let count = titlePane.find('#count-cell');
		let future = titlePane.find('#daysFuture-cell');
		let past = titlePane.find('#daysPast-cell');
		let nameField = optionsPane.find('#nameField-cell');
		let titleField = optionsPane.find('#titleField-cell');
		let birthdayField = optionsPane.find('#birthdayField-cell');

		window.onerror = (msg, url, lineNumber, columnNumber, error) => {
			console.log(msg, url, lineNumber, columnNumber, error);
		};

		bgColorPane.find('#position-cell').onChanged(value => {
			if (value.indexOf(',') !== -1) {
				draftBirthday.position = value.split(',');
			} else {
				draftBirthday.position = parseInt(value);
			}
		});

		let rowCell = bgColorPane.find('#backgroundColor-cell').parentNode;
		this.pickColor({ parent: rowCell, cell: rowCell.find('#backgroundColor-cell') }, (backgroundColor) => {
			if (draftBirthday.position) {
				if (typeof draftBirthday.position === 'object') {
					for (let each of draftBirthday.position) {
						birthdayListRow[parseInt(each.trim()) - 1].css({ backgroundColor });
					}
				} else {
					birthdayListRow[draftBirthday.position - 1].css({ backgroundColor });
				}
			}
			rowCell.find('#backgroundColor-cell').value = backgroundColor;
			rowCell.find('#backgroundColor-cell').setAttribute('value', backgroundColor);
		});

		let toggleDisplay = headerPane.find('#toggle-cell');
		toggleDisplay.addEventListener('change', e => {
			switch (toggleDisplay.value.toLowerCase()) {
				case 'show':
					draftDom.find('.crater-birthday-header').style.display = 'flex';
					break;
				case 'hide':
					draftDom.find('.crater-birthday-header').style.display = 'none';
					break;
			}
		});

		nameField.addEventListener('change', event => {
			switch (nameField.value.toLowerCase()) {
				case 'id':
					for (let person of draftBirthday.users) {
						if (!birthdayList.find('.no-users')) {
							birthdayListRow.forEach(div => {
								if (div.getAttribute('divid') === person.id) {
									div.find('.personName').textContent = person.id;
								}
							});
						}
					}
					break;
				case 'fullname':
					for (let person of draftBirthday.users) {
						if (!birthdayList.find('.no-users')) {
							birthdayListRow.forEach(div => {
								if (div.getAttribute('divid') === person.id) {
									div.find('.personName').textContent = person.displayName;
								}
							});
						}
					}
					break;
				case 'firstname':
					for (let person of draftBirthday.users) {
						if (!birthdayList.find('.no-users')) {
							birthdayListRow.forEach(div => {
								if (div.getAttribute('divid') === person.id) {
									div.find('.personName').textContent = person.firstName;
								}
							});
						}
					}
					break;
				case 'lastname':
					for (let person of draftBirthday.users) {
						if (!birthdayList.find('.no-users')) {
							birthdayListRow.forEach(div => {
								if (div.getAttribute('divid') === person.id) {
									div.find('.personName').textContent = person.lastName;
								}
							});
						}
					}
					break;
				case 'mail':
					for (let person of draftBirthday.users) {
						if (!birthdayList.find('.no-users')) {
							birthdayListRow.forEach(div => {
								if (div.getAttribute('divid') === person.id) {
									div.find('.personName').textContent = person.mail;
								}
							});
						}
					}
					break;
				case 'title':
					for (let person of draftBirthday.users) {
						if (!birthdayList.find('.no-users')) {
							birthdayListRow.forEach(div => {
								if (div.getAttribute('divid') === person.id) {
									div.find('.personName').textContent = person.Title;
								}
							});
						}
					}
					break;
				case 'phone':
					for (let person of draftBirthday.users) {
						if (!birthdayList.find('.no-users')) {
							birthdayListRow.forEach(div => {
								if (div.getAttribute('divid') === person.id) {
									div.find('.personName').textContent = person.phone;
								}
							});
						}
					}
					break;
				case 'birthday/anniversary':
					for (let person of draftBirthday.users) {
						if (!birthdayList.find('.no-users')) {
							birthdayListRow.forEach(div => {
								if (div.getAttribute('divid') === person.id) {
									if (draftBirthday.mode.toLowerCase() === "birthday") {
										div.find('.personName').textContent = new Date(person.birthDate).toString().split(`${new Date(person.birthDate).getFullYear()}`)[0];
									} else {
										div.find('.personName').textContent = person.anniversary;
									}
								}
							});
						}
					}
					break;
			}
		});


		titleField.addEventListener('change', event => {
			switch (titleField.value.toLowerCase()) {
				case 'id':
					for (let person of draftBirthday.users) {
						if (!birthdayList.find('.no-users')) {
							birthdayListRow.forEach(div => {
								if (div.getAttribute('divid') === person.id) {
									div.find('.title').textContent = person.id;
								}
							});
						}
					}
					break;
				case 'fullname':
					for (let person of draftBirthday.users) {
						if (!birthdayList.find('.no-users')) {
							birthdayListRow.forEach(div => {
								if (div.getAttribute('divid') === person.id) {
									div.find('.title').textContent = person.displayName;
								}
							});
						}
					}
					break;
				case 'firstname':
					for (let person of draftBirthday.users) {
						if (!birthdayList.find('.no-users')) {
							birthdayListRow.forEach(div => {
								if (div.getAttribute('divid') === person.id) {
									div.find('.title').textContent = person.firstName;
								}
							});
						}
					}
					break;
				case 'lastname':
					for (let person of draftBirthday.users) {
						if (!birthdayList.find('.no-users')) {
							birthdayListRow.forEach(div => {
								if (div.getAttribute('divid') === person.id) {
									div.find('.title').textContent = person.lastName;
								}
							});
						}
					}
					break;
				case 'mail':
					for (let person of draftBirthday.users) {
						if (!birthdayList.find('.no-users')) {
							birthdayListRow.forEach(div => {
								if (div.getAttribute('divid') === person.id) {
									div.find('.title').textContent = person.mail;
								}
							});
						}
					}
					break;
				case 'title':
					for (let person of draftBirthday.users) {
						if (!birthdayList.find('.no-users')) {
							birthdayListRow.forEach(div => {
								if (div.getAttribute('divid') === person.id) {
									div.find('.title').textContent = person.Title;
								}
							});
						}
					}
					break;
				case 'phone':
					for (let person of draftBirthday.users) {
						if (!birthdayList.find('.no-users')) {
							birthdayListRow.forEach(div => {
								if (div.getAttribute('divid') === person.id) {
									div.find('.title').textContent = person.phone;
								}
							});
						}
					}
					break;
				case 'birthday/anniversary':
					for (let person of draftBirthday.users) {
						if (!birthdayList.find('.no-users')) {
							birthdayListRow.forEach(div => {
								if (div.getAttribute('divid') === person.id) {
									if (draftBirthday.mode.toLowerCase() === "birthday") {
										div.find('.title').textContent = new Date(person.birthDate).toString().split(`${new Date(person.birthDate).getFullYear()}`)[0];
									} else {
										div.find('.title').textContent = person.anniversary;
									}
								}
							});
						}
					}
					break;
			}
		});


		birthdayField.addEventListener('change', event => {
			switch (birthdayField.value.toLowerCase()) {
				case 'id':
					for (let person of draftBirthday.users) {
						if (!birthdayList.find('.no-users')) {
							birthdayListRow.forEach(div => {

								if (div.getAttribute('divid') === person.id) {
									div.find('.birthDate').textContent = person.id;
								}
							});
						}
					}
					break;
				case 'fullname':
					for (let person of draftBirthday.users) {
						if (!birthdayList.find('.no-users')) {
							birthdayListRow.forEach(div => {
								if (div.getAttribute('divid') === person.id) {
									div.find('.birthDate').textContent = person.displayName;
								}
							});
						}
					}
					break;
				case 'firstname':
					for (let person of draftBirthday.users) {
						if (!birthdayList.find('.no-users')) {
							birthdayListRow.forEach(div => {
								if (div.getAttribute('divid') === person.id) {
									div.find('.birthDate').textContent = person.firstName;
								}
							});
						}
					}
					break;
				case 'lastname':
					for (let person of draftBirthday.users) {
						if (!birthdayList.find('.no-users')) {
							birthdayListRow.forEach(div => {
								if (div.getAttribute('divid') === person.id) {
									div.find('.birthDate').textContent = person.lastName;
								}
							});
						}
					}
					break;
				case 'mail':
					for (let person of draftBirthday.users) {
						if (!birthdayList.find('.no-users')) {
							birthdayListRow.forEach(div => {
								if (div.getAttribute('divid') === person.id) {
									div.find('.birthDate').textContent = person.mail;
								}
							});
						}
					}
					break;
				case 'title':
					for (let person of draftBirthday.users) {
						if (!birthdayList.find('.no-users')) {
							birthdayListRow.forEach(div => {
								if (div.getAttribute('divid') === person.id) {
									div.find('.birthDate').textContent = person.Title;
								}
							});
						}
					}
					break;
				case 'phone':
					for (let person of draftBirthday.users) {
						if (!birthdayList.find('.no-users')) {
							birthdayListRow.forEach(div => {
								if (div.getAttribute('divid') === person.id) {
									div.find('.birthDate').textContent = person.phone;
								}
							});
						}
					}
					break;
				case 'birthday/anniversary':
					for (let person of draftBirthday.users) {
						if (!birthdayList.find('.no-users')) {
							birthdayListRow.forEach(div => {
								if (div.getAttribute('divid') === person.id) {
									if (draftBirthday.mode.toLowerCase() === "birthday") {
										div.find('.birthDate').textContent = new Date(person.birthDate).toString().split(`${new Date(person.birthDate).getFullYear()}`)[0];
									} else {
										div.find('.birthDate').textContent = person.anniversary;
									}
								}
							});
						}
					}
					break;
			}
		});

		headerPane.find('#text-cell').onChanged(value => {
			draftDom.find('.crater-birthday-header-title').textContent = value;
		});

		let nameColorCell = namePane.find('#color-cell').parentNode;
		this.pickColor({ parent: nameColorCell, cell: nameColorCell.find('#color-cell') }, (color) => {
			draftDom.findAll('.personName').forEach(element => {
				element.css({ color });
			});
			nameColorCell.find('#color-cell').value = color;
			nameColorCell.find('#color-cell').setAttribute('value', color);
		});

		let headerCellParent = headerPane.find('#Image-cell').parentNode;
		this.uploadImage({ parent: headerCellParent }, (image) => {
			headerCellParent.find('#Image-cell').src = image.src;
			draftDom.find('.crater-birthday-header-img').src = image.src;
		});

		let headerColorCell = headerPane.find('#color-cell').parentNode;
		this.pickColor({ parent: headerColorCell, cell: headerColorCell.find('#color-cell') }, (color) => {
			draftDom.find('.crater-birthday-header-title').css({ color });
			headerColorCell.find('#color-cell').value = color;
			headerColorCell.find('#color-cell').setAttribute('value', color);
		});

		let headerBackgroundColorCell = headerPane.find('#background-color-cell').parentNode;
		this.pickColor({ parent: headerBackgroundColorCell, cell: headerBackgroundColorCell.find('#background-color-cell') }, (backgroundColor) => {
			draftDom.find('.crater-birthday-header').css({ backgroundColor });
			headerBackgroundColorCell.find('#background-color-cell').value = backgroundColor;
			headerBackgroundColorCell.find('#background-color-cell').setAttribute('value', backgroundColor);
		});

		let jobColorCell = jobPane.find('#color-cell').parentNode;
		this.pickColor({ parent: jobColorCell, cell: jobColorCell.find('#color-cell') }, (color) => {
			draftDom.findAll('.title').forEach(element => {
				element.css({ color });
			});
			jobColorCell.find('#color-cell').value = color;
			jobColorCell.find('#color-cell').setAttribute('value', color);
		});

		let birthdayColorCell = birthdayPane.find('#color-cell').parentNode;
		this.pickColor({ parent: birthdayColorCell, cell: birthdayColorCell.find('#color-cell') }, (color) => {
			draftDom.findAll('.birthDate').forEach(element => {
				element.css({ color });
			});
			birthdayColorCell.find('#color-cell').value = color;
			birthdayColorCell.find('#color-cell').setAttribute('value', color);
		});

		namePane.find('#fontSize-cell').onChanged(value => {
			draftDom.findAll('.personName').forEach(element => {
				element.css({ fontSize: value });
			});
		});

		jobPane.find('#fontSize-cell').onChanged(value => {
			draftDom.findAll('.title').forEach(element => {
				element.css({ fontSize: value });
			});
		});

		sectionPane.find('#section-height-cell').onChanged(value => {
			draftDom.find('.crater-birthday-div').css({ maxHeight: value });
		});

		headerPane.find('#header-height-cell').onChanged(value => {
			draftDom.find('.crater-birthday-header').css({ height: value });
			draftDom.find('.crater-birthday-header-img').css({ height: value });
			draftDom.find('.crater-birthday-header-img').css({ width: value });
		});

		birthdayPane.find('#fontSize-cell').onChanged(value => {
			draftDom.findAll('.birthDate').forEach(element => {
				element.css({ fontSize: value });
			});
		});

		namePane.find('#fontFamily-cell').onChanged(value => {
			draftDom.findAll('.personName').forEach(element => {
				element.css({ fontFamily: value });
			});
		});

		headerPane.find('#fontFamily-cell').onChanged(value => {
			draftDom.find('.crater-birthday-header-title').css({ fontFamily: value });
		});

		headerPane.find('#fontSize-cell').onChanged(value => {
			draftDom.find('.crater-birthday-header-title').css({ fontSize: value });
		});

		jobPane.find('#fontFamily-cell').onChanged(value => {
			draftDom.findAll('.title').forEach(element => {
				element.css({ fontFamily: value });
			});
		});

		birthdayPane.find('#fontFamily-cell').onChanged(value => {
			draftDom.findAll('.birthDate').forEach(element => {
				element.css({ fontFamily: value });
			});
		});

		modeCell.addEventListener('change', event => {
			let modeDiv = draftDom.findAll('.birthday');
			switch (modeCell.value.toLowerCase()) {
				case 'birthday':
					draftBirthday.mode = modeCell.value;
					for (let person of draftBirthday.users) {
						for (let position = 0; position < draftBirthday.users.length; position++) {
							if ((modeDiv[position]) && (!modeDiv[position].find('.no-users'))) {
								if (modeDiv[position].getAttribute('divid') === person.id) {
									modeDiv[position].find('.birthDate').textContent = new Date(person.birthDate).toString().split(`${new Date(person.birthDate).getFullYear()}`)[0];
								}
							}
						}
					}

					if (!titlePane.find('.title-settings').find('#Interval-cell').readOnly) {
						draftBirthday.interval = 999;
						titlePane.find('.title-settings').find('#Interval-cell').value = draftBirthday.interval;
						titlePane.find('.title-settings').find('#Interval-cell').readOnly = true;
					}
					break;
				case 'anniversary':
					draftBirthday.mode = modeCell.value;
					for (let person of draftBirthday.users) {
						for (let position = 0; position < draftBirthday.users.length; position++) {
							if ((modeDiv[position]) && (!modeDiv[position].find('.no-users'))) {
								if (modeDiv[position].getAttribute('divid') === person.id) {
									modeDiv[position].find('.birthDate').textContent = person.anniversary;
								}
							}
						}
					}
					console.log(draftBirthday.users);

					if (titlePane.find('.title-settings').find('#Interval-cell').readOnly) {
						titlePane.find('.title-settings').find('#Interval-cell').value = draftBirthday.interval;
						titlePane.find('.title-settings').find('#Interval-cell').readOnly = false;
					}

					break;
				default:
					draftBirthday.mode = 'Birthday';
			}
		});

		titlePane.find('.title-settings').find('#Interval-cell').onChanged(value => {
			draftBirthday.interval = parseInt(value);
		});

		sortCell.addEventListener('change', event => {
			switch (sortCell.value.toLowerCase()) {
				case 'name':
					draftBirthday.sortBy = sortCell.value;
					break;
				case 'birthday':
					draftBirthday.sortBy = sortCell.value;
					break;
				case 'anniversary':
					draftBirthday.sortBy = sortCell.value;
					break;
				default:
					draftBirthday.sortBy = 'birthday';
			}
		});

		orderCell.addEventListener('change', event => {
			switch (orderCell.value.toLowerCase()) {
				case 'ascending':

					let byDate = (draftBirthday.sortUsers.length !== 0) ? draftBirthday.sortUsers.slice(0) : draftBirthday.users.slice(0);

					byDate.sort((a, b) => {
						if (draftBirthday.sortBy.toLowerCase() === 'name') {
							if (a.firstName < b.firstName) return -1;
							if (b.firstName < a.firstName) return 1;
							return 0;
						}

						if (draftBirthday.sortBy.toLowerCase() === 'birthday') {
							if (new Date(a.birthDate) < new Date(b.birthDate)) return -1;
							if (new Date(b.birthDate) < new Date(a.birthDate)) return 1;
							return 0;
						}

						if (draftBirthday.sortBy.toLowerCase() === 'anniversary') {
							if (new Date(a.hireDate) > new Date(b.hireDate)) return -1;
							if (new Date(b.hireDate) > new Date(a.hireDate)) return 1;
							return 0;
						}
					});

					let craterBirthdayDiv = draftDom.find('.crater-birthday-div');
					let birthdays = craterBirthdayDiv.findAll('.birthday') as any;
					craterBirthdayDiv.innerHTML = '';
					const counted = (draftBirthday.count) ? draftBirthday.count : byDate.length;

					if (draftBirthday.mode.toLowerCase() === 'anniversary') {
						if (draftBirthday.interval == 999) {
							for (let i = 0; i < counted; i++) {
								const phoneExists = (byDate[i].phone) ? `<a href="tel:${byDate[i].phone}">
							<img class="img" src="http://www.free-icons-download.net/images/call-icon-16876.png"
								alt=""></a>
								<a href="sms:${byDate[i].phone}">
							<img class="img" src="http://www.free-icons-download.net/images/bubble-message-icon-74535.png"
								alt=""></a>
								` : '';
								let craterBirthdayHTML = `
							<div class="birthday" divid="${byDate[i].id}">
								<div class="birthday-image">
									<img class="name-image1"
										src=${byDate[i].image}
										alt="">
								</div>
								<div class="div">
									<div class="name"><span class="personName">${byDate[i].displayName}</span>
										<img class="name-image" src="https://img.icons8.com/material-sharp/24/000000/birthday.png"
											alt=""><br>
									</div>
									<span class="title">${byDate[i].mail}</span>
									<div class="date"><span class="birthDate">${byDate[i].anniversary}</span>
										<a href="mailto:${byDate[i].mail}?Subject=Happy%20Birthday%20${byDate[i].firstName}">
											<img class="img" src="https://img.icons8.com/material-two-tone/24/000000/composing-mail.png"
												alt=""></a>
									${phoneExists}
									</div>
								</div>
							</div>`;
								craterBirthdayDiv.innerHTML += craterBirthdayHTML;
							}
						} else {
							for (let i = 0; i < byDate.length; i++) {
								const phoneExists = (byDate[i].phone) ? `<a href="tel:${byDate[i].phone}">
							<img class="img" src="http://www.free-icons-download.net/images/call-icon-16876.png"
								alt=""></a>
								<a href="sms:${byDate[i].phone}">
							<img class="img" src="http://www.free-icons-download.net/images/bubble-message-icon-74535.png"
								alt=""></a>
								` : '';
								if (byDate[i].year === draftBirthday.interval) {
									console.log('retrieving users where year = ' + draftBirthday.interval + 'inside the code now ascending');
									let craterBirthdayHTML = `
								<div class="birthday" divid="${byDate[i].id}">
									<div class="birthday-image">
										<img class="name-image1"
											src=${byDate[i].image}
											alt="">
									</div>
									<div class="div">
										<div class="name"><span class="personName">${byDate[i].displayName}</span>
											<img class="name-image" src="https://img.icons8.com/material-sharp/24/000000/birthday.png"
												alt=""><br>
										</div>
										<span class="title">${byDate[i].mail}</span>
										<div class="date"><span class="birthDate">${byDate[i].anniversary}</span>
											<a href="mailto:${byDate[i].mail}?Subject=Happy%20Birthday%20${byDate[i].firstName}">
												<img class="img" src="https://img.icons8.com/material-two-tone/24/000000/composing-mail.png"
													alt=""></a>
									${phoneExists}
										</div>
									</div>
								</div>`;

									craterBirthdayDiv.innerHTML += craterBirthdayHTML;
								}
							}
						}
						if (craterBirthdayDiv.innerHTML.length === 0) {
							let craterBirthdayHTML = `
								<div style="padding:.5em;" class="no-users birthday">
									<p>Sorry, No users found</p>
								</div>`;

							craterBirthdayDiv.innerHTML = craterBirthdayHTML;
						}
					} else if (draftBirthday.mode.toLowerCase() === 'birthday') {
						for (let i = 0; i < counted; i++) {
							const phoneExists = (byDate[i].phone) ? `<a href="tel:${byDate[i].phone}">
							<img class="img" src="http://www.free-icons-download.net/images/call-icon-16876.png"
								alt=""></a>
								<a href="sms:${byDate[i].phone}">
							<img class="img" src="http://www.free-icons-download.net/images/bubble-message-icon-74535.png"
								alt=""></a>
								` : '';
							let craterBirthdayHTML = `<div class="birthday" divid="${byDate[i].id}">
							<div class="birthday-image">
								<img class="name-image1"
									src=${byDate[i].image}
									alt="">
							</div>
							<div class="div">
								<div class="name"><span class="personName">${byDate[i].displayName}</span>
									<img class="name-image" src="https://img.icons8.com/material-sharp/24/000000/birthday.png"
										alt=""><br>
								</div>
								<span class="title">${byDate[i].mail}</span>
								<div class="date"><span class="birthDate">${new Date(byDate[i].birthDate).toString().split(`${new Date(byDate[i].birthDate).getFullYear()}`)[0]}</span>
									<a href="mailto:${byDate[i].mail}?Subject=Happy%20Birthday%20${byDate[i].firstName}">
										<img class="img" src="https://img.icons8.com/material-two-tone/24/000000/composing-mail.png"
											alt=""></a>
									${phoneExists}
								</div>
							</div>
						</div>`;
							craterBirthdayDiv.innerHTML += craterBirthdayHTML;
						}
					}
					break;

				case 'descending':
					let byDate2 = (draftBirthday.sortUsers.length !== 0) ? draftBirthday.sortUsers.slice(0) : draftBirthday.users.slice(0);

					byDate2.sort((a, b) => {
						if (draftBirthday.sortBy.toLowerCase() === 'name') {
							if (a.firstName > b.firstName) return -1;
							if (b.firstName > a.firstName) return 1;
							return 0;
						}

						if (draftBirthday.sortBy.toLowerCase() === 'birthday') {
							if (new Date(a.birthDate) > new Date(b.birthDate)) return -1;
							if (new Date(b.birthDate) > new Date(a.birthDate)) return 1;
							return 0;
						}

						if (draftBirthday.sortBy.toLowerCase() === 'anniversary') {
							if (new Date(a.hireDate) < new Date(b.hireDate)) return -1;
							if (new Date(b.hireDate) < new Date(a.hireDate)) return 1;
							return 0;
						}
					});

					let craterBirthdayDiv2 = draftDom.find('.crater-birthday-div');

					craterBirthdayDiv2.innerHTML = '';
					const anotherCount = (draftBirthday.count) ? draftBirthday.count : byDate2.length;

					if (draftBirthday.mode.toLowerCase() === 'anniversary') {
						if (draftBirthday.interval == 999) {
							console.log('mode: ' + draftBirthday.mode, 'descending', draftBirthday.interval);
							for (let i = 0; i < anotherCount; i++) {
								const phoneExists = (byDate2[i].phone) ? `<a href="tel:${byDate2[i].phone}">
							<img class="img" src="http://www.free-icons-download.net/images/call-icon-16876.png"
								alt=""></a>
								<a href="sms:${byDate2[i].phone}">
							<img class="img" src="http://www.free-icons-download.net/images/bubble-message-icon-74535.png"
								alt=""></a>
								` : '';
								let craterBirthdayHTML2 = `<div class="birthday" divid="${byDate2[i].id}">
								<div class="birthday-image">
									<img class="name-image1"
										src=${byDate2[i].image}
										alt="">
								</div>
								<div class="div">
									<div class="name"><span class="personName">${byDate2[i].displayName}</span>
										<img class="name-image" src="https://img.icons8.com/material-sharp/24/000000/birthday.png"
											alt=""><br>
									</div>
									<span class="title">${byDate2[i].mail}</span>
									<div class="date"><span class="birthDate">${byDate2[i].anniversary}</span>
										<a href="mailto:${byDate2[i].mail}?Subject=Happy%20Birthday%20${byDate2[i].firstName}">
											<img class="img" src="https://img.icons8.com/material-two-tone/24/000000/composing-mail.png"
												alt=""></a>
									${phoneExists}
									</div>
								</div>
							</div>`;
								craterBirthdayDiv2.innerHTML += craterBirthdayHTML2;
							}
						} else {
							for (let x = 0; x < anotherCount; x++) {
								if (byDate2[x].year === draftBirthday.interval) {
									const phoneExists = (byDate2[x].phone) ? `<a href="tel:${byDate2[x].phone}">
							<img class="img" src="http://www.free-icons-download.net/images/call-icon-16876.png"
								alt=""></a>
								<a href="sms:${byDate2[x].phone}">
							<img class="img" src="http://www.free-icons-download.net/images/bubble-message-icon-74535.png"
								alt=""></a>
								` : '';
									let craterBirthdayHTML2 = `<div class="birthday" divid="${byDate2[x].id}">
									<div class="birthday-image">
										<img class="name-image1"
											src=${byDate2[x].image}
											alt="">
									</div>
									<div class="div">
										<div class="name"><span class="personName">${byDate2[x].displayName}</span>
											<img class="name-image" src="https://img.icons8.com/material-sharp/24/000000/birthday.png"
												alt=""><br>
										</div>
										<span class="title">${byDate2[x].mail}</span>
										<div class="date"><span class="birthDate">${byDate2[x].anniversary}</span>
											<a href="mailto:${byDate2[x].mail}?Subject=Happy%20Birthday%20${byDate2[x].firstName}">
												<img class="img" src="https://img.icons8.com/material-two-tone/24/000000/composing-mail.png"
													alt=""></a>
										${phoneExists}
										</div>
									</div>
								</div>`;

									craterBirthdayDiv2.innerHTML += craterBirthdayHTML2;
								}
							}
						}
						if (craterBirthdayDiv2.innerHTML.length === 0) {
							let craterBirthdayHTML2 = `
								<div style="padding:.5em;" class="no-users birthday">
									<p>Sorry, No users found</p>
								</div>`;
							craterBirthdayDiv2.innerHTML = craterBirthdayHTML2;
						}

					} else {
						if (draftBirthday.mode.toLowerCase() === 'birthday') {
							for (let i = 0; i < anotherCount; i++) {
								const phoneExists = (byDate2[i].phone) ? `<a href="tel:${byDate2[i].phone}">
							<img class="img" src="http://www.free-icons-download.net/images/call-icon-16876.png"
								alt=""></a>
								<a href="sms:${byDate2[i].phone}">
							<img class="img" src="http://www.free-icons-download.net/images/bubble-message-icon-74535.png"
								alt=""></a>
								` : '';
								let craterBirthdayHTML2 = `<div class="birthday" divid="${byDate2[i].id}">
								<div class="birthday-image">
									<img class="name-image1"
										src=${byDate2[i].image}
										alt="">
								</div>
								<div class="div">
									<div class="name"><span class="personName">${byDate2[i].displayName}</span>
										<img class="name-image" src="https://img.icons8.com/material-sharp/24/000000/birthday.png"
											alt=""><br>
									</div>
									<span class="title">${byDate2[i].mail}</span>
									<div class="date"><span class="birthDate">${new Date(byDate2[i].birthDate).toString().split(`${new Date(byDate2[i].birthDate).getFullYear()}`)[0]}</span>
										<a href="mailto:${byDate2[i].mail}?Subject=Happy%20Birthday%20${byDate2[i].firstName}">
											<img class="img" src="https://img.icons8.com/material-two-tone/24/000000/composing-mail.png"
												alt=""></a>
									${phoneExists}
									</div>
								</div>
							</div>`;
								craterBirthdayDiv2.innerHTML += craterBirthdayHTML2;
							}
						}
					}

					break;
			}
		});

		count.onChanged(value => {
			if (typeof value === "number") draftBirthday.count = value;
			else if (typeof value === "string") draftBirthday.count = parseInt(value);
		});

		let addDays = (givenDate) => {
			let date = new Date(Number(new Date()));
			date.setDate(date.getDate() + parseInt(givenDate));
			return date;
		};

		let subtractDays = (givenDate) => {
			let date = new Date(Number(new Date()));
			date.setDate(date.getDate() - parseInt(givenDate));
			return date;
		};

		let sortArrayAsc = (arrayToPush, array, numberCount) => {
			array.sort((a, b) => {
				if (a.newBirthDate < b.newBirthDate) return -1;
				if (b.newBirthDate < a.newBirthDate) return 1;
			});

			for (let i = 0; i < numberCount; i++) {
				if (arrayToPush.length === 0) {
					arrayToPush.push(array[i]);
				} else {
					if (arrayToPush.indexOf(array[i]) === -1) {
						arrayToPush.push(array[i]);
					}
				}
			}
		};

		let sortArrayDes = (arrayToPush, array, numberCount) => {
			array.sort((a, b) => {
				if (a.newBirthDate > b.newBirthDate) return -1;
				if (b.newBirthDate > a.newBirthDate) return 1;
			});

			for (let i = 0; i < numberCount; i++) {
				if (arrayToPush.length === 0) {
					arrayToPush.push(array[i]);
				} else {
					if (arrayToPush.indexOf(array[i]) === -1) {
						arrayToPush.push(array[i]);
					}
				}
			}
		};

		future.addEventListener('change', e => {
			let value = future.value;

			if (draftBirthday.mode.toLowerCase() === 'birthday') {
				draftBirthday.beforeYear = parseInt(value);
				let futureSort = [];
				for (let z = 0; z < draftBirthday.users.length; z++) {
					let birthMonth = new Date(draftBirthday.users[z].birthDate).getMonth() + 1;
					let userBirthDate = new Date(draftBirthday.users[z].birthDate).getDate();
					let newBirthDate = new Date(`${birthMonth}/${userBirthDate}/${new Date().getFullYear()}`);
					draftBirthday.users[z].newBirthDate = newBirthDate;

					if ((draftBirthday.users[z].newBirthDate <= new Date(`12/31/${new Date().getFullYear()}`)) && (draftBirthday.users[z].newBirthDate >= new Date())) {
						if (futureSort.length === 0) {
							futureSort.push(draftBirthday.users[z]);
						} else {
							if (futureSort.indexOf(draftBirthday.users[z]) === -1) {
								futureSort.push(draftBirthday.users[z]);
							}
						}
					}
				}

				sortArrayAsc(draftBirthday.sortUsers, futureSort, parseInt(value));
			}

			if (draftBirthday.mode.toLowerCase() === 'anniversary') {
				draftBirthday.beforeYear = parseInt(value);
				let futureSort = [];
				for (let z = 0; z < draftBirthday.users.length; z++) {
					let date = new Date();
					//@ts-ignore
					let personHireDate = new Date(draftBirthday.users[z].hireDate);
					let timeLeft = date.getTime() - personHireDate.getTime();

					let daysLeft = Math.floor(timeLeft / (1000 * 60 * 60 * 24));
					let hoursLeft = Math.floor((timeLeft % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
					let minutesLeft = Math.floor((timeLeft % (1000 * 60 * 60)) / (1000 * 60));
					let secondsLeft = Math.floor((timeLeft % (1000 * 60)) / (1000));

					draftBirthday.users[z].newBirthDate = this.getDAYS(daysLeft).year;
					if (this.getDAYS(daysLeft).month >= 5) {
						if (futureSort.length === 0) {
							futureSort.push(draftBirthday.users[z]);
						} else {
							if (futureSort.indexOf(draftBirthday.users[z]) === -1) {
								futureSort.push(draftBirthday.users[z]);
							}
						}
					}
				}

				sortArrayDes(draftBirthday.sortUsers, futureSort, parseInt(value));
			}
		});

		past.addEventListener('change', e => {
			let value = past.value;

			if (draftBirthday.mode.toLowerCase() === 'birthday') {
				draftBirthday.endAnniversaryAfter = parseInt(value);

				let pastSort = [];
				for (let z = 0; z < draftBirthday.users.length; z++) {
					let birthMonth = new Date(draftBirthday.users[z].birthDate).getMonth() + 1;
					let userBirthDate = new Date(draftBirthday.users[z].birthDate).getDate();
					let newBirthDate = new Date(`${birthMonth}/${userBirthDate}/${new Date().getFullYear()}`);
					draftBirthday.users[z].newBirthDate = newBirthDate;

					if ((draftBirthday.users[z].newBirthDate <= new Date()) && (draftBirthday.users[z].newBirthDate >= new Date(`01/01/${new Date().getFullYear()}`))) {
						if (pastSort.length === 0) {
							pastSort.push(draftBirthday.users[z]);
						} else {
							if (pastSort.indexOf(draftBirthday.users[z]) === -1) {
								pastSort.push(draftBirthday.users[z]);
							}
						}
					}

					if (pastSort.length === 0) {
						if ((draftBirthday.users[z].newBirthDate <= new Date()) && (draftBirthday.users[z].newBirthDate >= new Date(`01/01/${new Date().getFullYear() - 1}`))) {
							if (pastSort.length === 0) {
								pastSort.push(draftBirthday.users[z]);
							} else {
								if (pastSort.indexOf(draftBirthday.users[z]) === -1) {
									pastSort.push(draftBirthday.users[z]);
								}
							}
						}
					}
				}

				sortArrayDes(draftBirthday.sortUsers, pastSort, parseInt(value));
			}

			if (draftBirthday.mode.toLowerCase() === 'anniversary') {
				draftBirthday.beforeYear = parseInt(value);
				let pastSort = [];
				for (let z = 0; z < draftBirthday.users.length; z++) {
					let date = new Date();
					//@ts-ignore
					let personHireDate = new Date(draftBirthday.users[z].hireDate);
					let timeLeft = date.getTime() - personHireDate.getTime();

					let daysLeft = Math.floor(timeLeft / (1000 * 60 * 60 * 24));
					let hoursLeft = Math.floor((timeLeft % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
					let minutesLeft = Math.floor((timeLeft % (1000 * 60 * 60)) / (1000 * 60));
					let secondsLeft = Math.floor((timeLeft % (1000 * 60)) / (1000));

					draftBirthday.users[z].newBirthDate = this.getDAYS(daysLeft).year;
					if (this.getDAYS(daysLeft).month < 5) {
						if (pastSort.length === 0) {
							pastSort.push(draftBirthday.users[z]);
						} else {
							if (pastSort.indexOf(draftBirthday.users[z]) === -1) {
								pastSort.push(draftBirthday.users[z]);
							}
						}
					}
				}

				sortArrayDes(draftBirthday.sortUsers, pastSort, parseInt(value));
			}
		});

		this.paneContent.find('.update-pane').find('#fetch-birthday').addEventListener('click', event => {
			try {
				draftBirthday.fetched = false;
				this.paneContent.find('#update-message').textContent = 'Fetching directory now... Please wait';
				this.paneContent.find('#fetch-birthday').render({
					element: 'img', attributes: { class: 'birthday-loading-image crater-icon', src: this.sharePoint.images.loading, style: { width: '20px', height: '20px' } }
				});
				setTimeout(() => {
					if (draftBirthday.fetched) {
						this.paneContent.find('#update-message').textContent = 'Directory Updated! Sort the list to view it';
						this.paneContent.find('#fetch-birthday').innerHTML = 'UPDATED';
					}
				}, 7000);
				this.getUser();
			} catch (error) {
				console.log(error.message);
			}
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').find('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = draftDom.innerHTML;
			this.element.css(draftDom.css());
			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;
		});
	}
}

class Twitter extends CraterWebParts {
	public key: any;
	public params: any;
	public elementModifier: any = new ElementModifier();
	public paneContent: any;
	public element: any;

	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.params = params;
	}

	public render(params) {
		let twitter = this.createKeyedElement({
			element: 'div', attributes: { class: 'crater-twitter crater-component', 'data-type': 'twitter' }, children: [
				{ element: 'div', attributes: { class: 'crater-twitter-feed' } }
			]
		});


		this.key = twitter.dataset.key;

		this.sharePoint.properties.pane.content[this.key].settings.twitter = {
			'data-width': '1000',
			'data-height': '500',
			'data-link-color': '#d6ff27',
			'data-theme': 'light',
			'data-tweet-limit': '3',
			username: 'TwitterDev',
		};


		return twitter;
	}

	public rendered(params) {
		this.element = params.element;
		this.key = params.element.dataset.key;

		this.runTwitter();
	}

	public runTwitter() {
		let twitterObj = this.sharePoint.properties.pane.content[this.key].settings.twitter;
		let self = this;

		if (twitterObj['data-tweet-limit']) {
			this.element.style.maxHeight = twitterObj['data-height'];
		} else {
			this.element.style.maxHeight = '100%';
		}

		this.element.querySelector('.crater-twitter-feed').innerHTML = `<a class='twitter-timeline' id='twitter-embed' max-width='100%' data-width= ${twitterObj['data-width']} data-height=${twitterObj['data-height']} lang='EN' href='https://twitter.com/${twitterObj.username}' data-link-color=${twitterObj['data-link-color']} data-theme=${twitterObj['data-theme']} data-tweet-limit=${twitterObj['data-tweet-limit']}>
		Tweets by @${twitterObj.username}</a>`;

		let twitterFunction = (s: string, id: string) => {
			let js: any;
			if (self.element.querySelector('#twitter-wjs')) self.element.querySelector('#twitter-wjs').remove();
			let t = window['twttr'] || {};
			js = document.createElement(s);
			js.id = id;
			js.src = "https://platform.twitter.com/widgets.js";
			self.element.querySelector('.crater-twitter-feed').parentNode.insertBefore(js, self.element.querySelector('.crater-twitter-feed'));

			t._e = [];
			t.ready = (f) => {
				t._e.push(f);
			};

			return t;
		};

		window['twttr'] = (twitterFunction("script", "twitter-wjs"));

		// Define our custom event handlers
		let clickEventToAnalytics = (intentEvent) => {
			if (!intentEvent) return;
			var label = intentEvent.region;
			//@ts-ignore
			pageTracker._trackEvent('twitter_web_intents', intentEvent.type, label);
		};

		let tweetIntentToAnalytics = (intentEvent) => {
			if (!intentEvent) return;
			var label = "tweet";
			//@ts-ignore
			pageTracker._trackEvent(
				'twitter_web_intents',
				intentEvent.type,
				label
			);
		};

		let likeIntentToAnalytics = (intentEvent) => {
			tweetIntentToAnalytics(intentEvent);
		};

		let retweetIntentToAnalytics = (intentEvent) => {
			if (!intentEvent) return;
			var label = intentEvent.data.source_tweet_id;
			//@ts-ignore
			pageTracker._trackEvent(
				'twitter_web_intents',
				intentEvent.type,
				label
			);
		};

		let followIntentToAnalytics = (intentEvent) => {
			if (!intentEvent) return;
			var label = intentEvent.data.user_id + " (" + intentEvent.data.screen_name + ")";
			//@ts-ignore
			pageTracker._trackEvent(
				'twitter_web_intents',
				intentEvent.type,
				label
			);
		};

		// Wait for the asynchronous resources to load
		try {
			//@ts-ignore
			twttr.ready((twttr) => {
				// Now bind our custom intent events
				twttr.events.bind('click', clickEventToAnalytics);
				twttr.events.bind('tweet', tweetIntentToAnalytics);
				twttr.events.bind('retweet', retweetIntentToAnalytics);
				twttr.events.bind('like', likeIntentToAnalytics);
				twttr.events.bind('follow', followIntentToAnalytics);
			});
		} catch (error) {
			console.log(error.message);
		}
	}

	public setUpPaneContent(params) {
		let key = params.element.dataset['key'];
		this.element = this.sharePoint.properties.pane.content[key].draft.dom;
		this.paneContent = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'crater-property-content' }
		}).monitor();
		let twitterObj = this.sharePoint.properties.pane.content[key].settings.twitter;

		if (this.sharePoint.properties.pane.content[key].draft.pane.content.length !== 0) {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[key].draft.pane.content;
		}
		else if (this.sharePoint.properties.pane.content[key].content.length !== 0) {
			this.paneContent.innerHTML = this.sharePoint.properties.pane.content[key].content;
		} else {

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'twitter-options-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'OPTIONS'
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'message-note' }, children: [
							{
								element: 'div', attributes: { class: 'message-text' }, children: [
									{ element: 'p', attributes: { style: { color: 'green' } }, text: `NOTE: Clear the 'number of tweets' field to display all tweets` },
									{ element: 'p', attributes: { style: { color: 'green' } }, text: `      Please Enter a Valid Username` }
								]
							}
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'username', value: twitterObj.username
							}),
							this.elementModifier.cell({
								element: 'select', name: 'theme', options: ['LIGHT', 'DARK'], value: twitterObj['data-theme'].toUpperCase()
							}),
							this.elementModifier.cell({
								element: 'input', name: 'link-color', value: twitterObj['data-link-color']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'number-of-tweets', value: twitterObj['data-tweet-limit']
							})
						]
					})
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'twitter-size-pane card' }, children: [
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'SIZE CONTROL'
							})
						]
					}),

					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'width', value: twitterObj['data-width']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'height', value: twitterObj['data-height']
							})
						]
					})
				]
			});

		}
		return this.paneContent;
	}

	public listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		this.paneContent = this.sharePoint.app.querySelector('.crater-property-content').monitor();
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;
		let twitterObj = this.sharePoint.properties.pane.content[this.key].settings.twitter;

		let twitterOptions = this.paneContent.querySelector('.twitter-options-pane');
		let twitterSize = this.paneContent.querySelector('.twitter-size-pane');

		twitterOptions.querySelector('#username-cell').onChanged(value => {
			twitterObj.username = value;
		});

		twitterOptions.querySelector('#theme-cell').onChanged(value => {
			twitterObj['data-theme'] = value.toLowerCase();
		});

		twitterOptions.querySelector('#number-of-tweets-cell').onChanged(value => {
			twitterObj['data-tweet-limit'] = value;
		});


		let linkColorCell = twitterOptions.querySelector('#link-color-cell').parentNode;
		this.pickColor({ parent: linkColorCell, cell: linkColorCell.querySelector('#link-color-cell') }, (color) => {
			twitterObj['data-link-color'] = ColorPicker.rgbToHex(color);
			linkColorCell.querySelector('#link-color-cell').value = ColorPicker.rgbToHex(color);
			linkColorCell.querySelector('#link-color-cell').setAttribute('value', ColorPicker.rgbToHex(color));
		});

		twitterSize.querySelector('#width-cell').onChanged(value => {
			twitterObj['data-width'] = value;
		});

		twitterSize.querySelector('#height-cell').onChanged(value => {
			twitterObj['data-height'] = value;
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = draftDom.innerHTML;
			this.element.css(draftDom.css());
			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;
		});
	}
}

{
	CraterWebParts.prototype['twitter'] = params => {
		let twitter = new Twitter({ sharePoint: params.sharePoint });
		return twitter[params.action](params);
	};

	CraterWebParts.prototype['birthday'] = params => {
		let birthday = new Birthday({ sharePoint: params.sharePoint });
		return birthday[params.action](params);
	};

	CraterWebParts.prototype['employeedirectory'] = params => {
		let employeeDirectory = new EmployeeDirectory({ sharePoint: params.sharePoint });
		return employeeDirectory[params.action](params);
	};

	CraterWebParts.prototype['power'] = params => {
		let power = new Power({ sharePoint: params.sharePoint });
		return power[params.action](params);
	};

	CraterWebParts.prototype['event'] = params => {
		let event = new Event({ sharePoint: params.sharePoint });
		return event[params.action](params);
	};

	CraterWebParts.prototype['youtube'] = params => {
		let youtube = new YouTube({ sharePoint: params.sharePoint });
		return youtube[params.action](params);
	};

	CraterWebParts.prototype['facebook'] = params => {
		let facebook = new Facebook({ sharePoint: params.sharePoint });
		return facebook[params.action](params);
	};

	CraterWebParts.prototype['beforeafter'] = params => {
		let beforeafter = new BeforeAfter({ sharePoint: params.sharePoint });
		return beforeafter[params.action](params);
	};

	CraterWebParts.prototype['map'] = params => {
		let map = new Map({ sharePoint: params.sharePoint });
		return map[params.action](params);
	};

	CraterWebParts.prototype['datelist'] = params => {
		let datelist = new DateList({ sharePoint: params.sharePoint });
		return datelist[params.action](params);
	};

	CraterWebParts.prototype['instagram'] = params => {
		let instagram = new Instagram({ sharePoint: params.sharePoint });
		return instagram[params.action](params);
	};

	CraterWebParts.prototype['carousel'] = params => {
		let carousel = new Carousel({ sharePoint: params.sharePoint });
		return carousel[params.action](params);
	};

	CraterWebParts.prototype['events'] = params => {
		let events = new Events({ sharePoint: params.sharePoint });
		return events[params.action](params);
	};

	CraterWebParts.prototype['tab'] = params => {
		let tab = new Tab({ sharePoint: params.sharePoint });
		return tab[params.action](params);
	};

	CraterWebParts.prototype['countdown'] = params => {
		let countdown = new CountDown({ sharePoint: params.sharePoint });
		return countdown[params.action](params);
	};

	CraterWebParts.prototype['button'] = params => {
		let button = new Button({ sharePoint: params.sharePoint });
		return button[params.action](params);
	};

	CraterWebParts.prototype['icons'] = params => {
		let icons = new Icons({ sharePoint: params.sharePoint });
		return icons[params.action](params);
	};

	CraterWebParts.prototype['textarea'] = params => {
		let textarea = new TextArea({ sharePoint: params.sharePoint });
		return textarea[params.action](params);
	};

	CraterWebParts.prototype['news'] = params => {
		let news = new News({ sharePoint: params.sharePoint });
		return news[params.action](params);
	};

	CraterWebParts.prototype['crater'] = params => {
		let crater = new Crater({ sharePoint: params.sharePoint });
		return crater[params.action](params);
	};

	CraterWebParts.prototype['panel'] = params => {
		let panel = new Panel({ sharePoint: params.sharePoint });
		return panel[params.action](params);
	};

	CraterWebParts.prototype['table'] = params => {
		let table = new Table({ sharePoint: params.sharePoint });
		return table[params.action](params);
	};

	CraterWebParts.prototype['section'] = params => {
		let section = new Section({ sharePoint: params.sharePoint });
		return section[params.action](params);
	};

	CraterWebParts.prototype['sample'] = params => {
		let sample = new Sample({ sharePoint: params.sharePoint });
	};

	CraterWebParts.prototype['slider'] = params => {
		let slider = new Slider({ sharePoint: params.sharePoint });
		return slider[params.action](params);
	};

	CraterWebParts.prototype['list'] = params => {
		let list = new List({ sharePoint: params.sharePoint });
		return list[params.action](params);
	};

	CraterWebParts.prototype['counter'] = params => {
		let counter = new Counter({ sharePoint: params.sharePoint });
		return counter[params.action](params);
	};

	CraterWebParts.prototype['tiles'] = params => {
		let tiles = new Tiles({ sharePoint: params.sharePoint });
		return tiles[params.action](params);
	};
}

export { CraterWebParts };