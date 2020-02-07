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
require('./../../../node_modules/froala-editor/css/froala_editor.pkgd.min.css');
const factory = require('./powerbi.js');

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

class CraterWebParts {
	public elementModifier = new ElementModifier();
	public sharePoint: any;
	public connectable: any = [
		'list', 'slider', 'counter', 'tiles', 'news', 'table', 'icons', 'button', 'events', 'carousel', 'datelist'
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
		params.parent.querySelectorAll('.pick-color').forEach(element => {
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
				params.parent.querySelectorAll('.crater-color-picker').forEach(element => {
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
			params.parent.querySelectorAll('.pick-color').forEach(element => {
				element.remove();
			});
		});
	}

	//set up the image uploader
	public uploadImage(params, callBack) {
		//remove all the uploader
		params.parent.querySelectorAll('.upload-form').forEach(element => {
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
				params.parent.querySelectorAll('.upload-form').forEach(element => {
					element.remove();
				});
				//add the uploader
				this.elementModifier.importImage({ parent: params.parent, name: 'upload', attributes: { class: 'upload-form' } }, (image) => {
					params.parent.querySelectorAll('.upload-form').forEach(element => {
						element.remove();
					});
					callBack(image);
				});
			});
		});

		params.parent.addEventListener('mouseleave', (event) => {
			params.parent.querySelectorAll('.upload-image').forEach(element => {
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
			paste: this.sharePoint.images.paste
		};

		for (let option of params.options) {
			optionContainer.makeElement({
				element: 'img', attributes: {
					class: 'webpart-option', id: option.toLowerCase() + '-me', src: options[option.toLowerCase()], alt: option, title: `${option} ${params.title}`
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
					webpart.querySelector('.webpart-options').querySelector('#paste-me').show();
				} else {
					webpart.querySelector('.webpart-options').querySelector('#paste-me').hide();
				}
			}
		};

		element.addEventListener('mouseenter', event => {
			if (element.hasAttribute('data-key') && this.sharePoint.inEditMode()) {
				element.querySelector('.webpart-options').show();
				handlePaste(element);
			}
		});

		element.addEventListener('mouseleave', event => {
			if (element.hasAttribute('data-key')) {
				element.querySelector('.webpart-options').hide();
			}
		});

		element.querySelectorAll('.keyed-element').forEach(keyedElement => {
			keyedElement.addEventListener('mouseenter', event => {
				if (keyedElement.hasAttribute('data-key') && this.sharePoint.inEditMode()) {
					keyedElement.querySelector('.webpart-options').show();
					handlePaste(keyedElement);
				}
			});

			keyedElement.addEventListener('mouseleave', event => {
				if (keyedElement.hasAttribute('data-key')) {
					keyedElement.querySelector('.webpart-options').hide();
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
				element.querySelector('.webpart-options').css({ top: 'unset' });
			}
			align[params.alignOptions] = '0px';
			element.querySelector('.webpart-options').css(align);
			if (params.alignOptions == 'center') {
				element.querySelector('.webpart-options').css({ margin: '0em 3em' });
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
		this.sharePoint.properties.pane.content[this.key].settings.employees = {};
		let settings = this.sharePoint.properties.pane.content[this.key].settings;

		settings.searchType = 'All';
		settings.searchQuery = '';
		settings.mailApp = 'Outlook';
		settings.messageApp = 'Teams';
		settings.callApp = 'Teams';

		localStorage[`crater-${this.key}`] = JSON.stringify(settings);
		return employeeDirectory;
	}

	public rendered(params) {
		this.sharePoint = params.sharePoint;
		this.element = params.element;
		this.key = this.element.dataset.key;
		let displayed = false;

		let settings = this.sharePoint.properties.pane.content[this.key].settings;

		let display = this.element.querySelector('.crater-employee-directory-display');

		let gmail = 'https://mail.google.com';
		let outlook = 'https://outlook.office365.com/mail';
		let yahoo = 'https://mail.yahoo.com';
		let skype = 'https://www.skype.com/en/business/';
		let teams = 'https://teams.microsoft.com/_#/conversations';

		let me;

		display.innerHTML = '';

		display.makeElement({ element: 'img', attributes: { src: this.sharePoint.images.loading, class: 'crater-icon', style: { alignSelf: 'center', justifySelf: 'center' } } });

		let getUsers = async () => {
			this.sharePoint.connection.getWithGraph().then(client => {
				client.api('/users')
					.select('mail, displayName, givenName, id, jobTitle, mobilePhone')
					.get((_error: any, _result: MicrosoftGraph.User, _rawResponse?: any) => {
						this.users = _result['value'];

						for (let employee of this.users) {
							settings.employees[employee.id] = employee;
							let getImage = () => {
								client.api(`/users/${employee.id}/photo/$value`)
									.responseType('blob')
									.get((error: any, result: any, rawResponse?: any) => {
										if (!func.setNotNull(result)) return;
										settings.employees[employee.id].photo = result;
										if (displayed) {
											display.querySelector(`#row-${employee.id}`).querySelector('.crater-employee-directory-dp').src = window.URL.createObjectURL(result);
										}
									});
							};

							let getDepartment = () => {
								client.api(`/users/${employee.id}/department`)
									.get((error: any, result: any, rawResponse?: any) => {
										if (!func.setNotNull(result)) return;
										settings.employees[employee.id].department = result.value;
									});
							};

							getImage();
							getDepartment();
						}

						this.displayUsers(display);
						displayed = true;
					});
			});
		};

		if (!this.sharePoint.isLocal) {
			getUsers();
		} else {
			let sample = { mail: 'kennedy.ikeka@ipigroupng.com', id: this.key, displayName: 'Ikeka Kennedy', jobTitle: 'Programmer' };
			this.users = [];
			settings.employees = {};
			for (let i = 0; i < 4; i++) {
				this.users.push(sample);
				settings.employees[this.key] = sample;
			}

			this.displayUsers(display);
		}

		let changeSearchType = this.element.querySelector('#crater-employee-directory-search-type');
		let changeSearchQuery = this.element.querySelector('#crater-employee-directory-search-query');
		let sync = this.element.querySelector('#crater-employee-directory-sync');

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

		let menu = this.element.querySelector('.crater-employee-directory-menu');
		if (menu.position().width < 400) {
			menu.css({ gridTemplateColumns: '1fr' });
		} else {
			menu.cssRemove(['grid-template-columns']);
		}

		display.addEventListener('click', event => {
			let element = event.target;

			if (element.classList.contains('crater-employee-directory-toggle-view')) {
				element.classList.toggle('open');
				let row = element.getParents('.crater-employee-directory-row');

				display.querySelectorAll('.crater-employee-directory-other-details').forEach(other => {
					other.remove();
				});

				if (element.classList.contains('open')) {
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

		for (let i in this.users) {
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
									{ element: 'img', attributes: { class: 'crater-employee-directory-icon crater-employee-directory-mail', src: this.sharePoint.images.mail } },
									{ element: 'img', attributes: { class: 'crater-employee-directory-icon crater-employee-directory-message', src: this.sharePoint.images.message } },
									{ element: 'img', attributes: { class: 'crater-employee-directory-icon crater-employee-directory-phone', src: this.sharePoint.images.phone } }
									// {
									// 	element: 'a', attributes: { href: `mailto:${employee.mail}`, id: 'crater-employee-directory-mail' }, children: [
									// 		{ element: 'img', attributes: { class: 'crater-employee-directory-icon', src: this.sharePoint.images.mail } },
									// 	]
									// },
									// {
									// 	element: 'a', attributes: { href: `sms:${employee.mail}`, id: 'crater-employee-directory-message' }, children: [
									// 		{ element: 'img', attributes: { class: 'crater-employee-directory-icon', src: this.sharePoint.images.message } },
									// 	]
									// },
									// {
									// 	element: 'a', attributes: { href: `callto:${employee.mail}`, id: 'crater-employee-directory-phone' }, children: [
									// 		{ element: 'img', attributes: { class: 'crater-employee-directory-icon', src: this.sharePoint.images.phone } },
									// 	]
									// }
								]
							}
						]
					},
					{ element: 'img', attributes: { class: 'crater-employee-directory-icon crater-employee-directory-toggle-view', src: this.sharePoint.images.append } }
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
		this.paneContent = this.sharePoint.app.querySelector('.crater-property-content').monitor();
		let settings = this.sharePoint.properties.pane.content[this.key].settings;

		let domDraft = this.sharePoint.properties.pane.content[this.key].draft.dom;

		let menuPane = this.paneContent.querySelector('.menu-pane');
		let searchTypePane = this.paneContent.querySelector('.search-type-pane');
		let displayPane = this.paneContent.querySelector('.display-pane');
		let appsPane = this.paneContent.querySelector('.apps-pane');

		appsPane.querySelector('#Mail-cell').onChanged();
		appsPane.querySelector('#Message-cell').onChanged();
		appsPane.querySelector('#Call-cell').onChanged();

		menuPane.querySelector('#Border-Color-cell').onChanged(borderColor => {
			let borderType = menuPane.querySelector('#Border-Type-cell').value || 'solid';
			let borderSize = menuPane.querySelector('#Border-Size-cell').value || '1px';
			domDraft.querySelector('.crater-employee-directory-menu').css({ border: `${borderSize} ${borderType} ${borderColor}` });
		});

		menuPane.querySelector('#Border-Size-cell').onChanged(borderSize => {
			let borderType = menuPane.querySelector('#Border-Type-cell').value || 'solid';
			let borderColor = menuPane.querySelector('#Border-Color-cell').value || '1px';
			domDraft.querySelector('.crater-employee-directory-menu').css({ border: `${borderSize} ${borderType} ${borderColor}` });
		});

		menuPane.querySelector('#Border-Type-cell').onChanged(borderType => {
			let borderSize = menuPane.querySelector('#Border-Size-cell').value || 'solid';
			let borderColor = menuPane.querySelector('#Border-Color-cell').value || '1px';
			domDraft.querySelector('.crater-employee-directory-menu').css({ border: `${borderSize} ${borderType} ${borderColor}` });
		});

		menuPane.querySelector('#Background-Color-cell').onChanged(backgroundColor => {
			domDraft.querySelector('.crater-employee-directory-menu').css({ backgroundColor });
		});

		searchTypePane.querySelector('#Shadow-cell').onChanged(boxShadow => {
			domDraft.querySelector('#crater-employee-directory-search-type').css({ boxShadow });
		});

		searchTypePane.querySelector('#Border-cell').onChanged(border => {
			domDraft.querySelector('#crater-employee-directory-search-type').css({ border });
		});

		searchTypePane.querySelector('#Color-cell').onChanged(color => {
			domDraft.querySelector('#crater-employee-directory-search-type').css({ color });
		});

		searchTypePane.querySelector('#Background-Color-cell').onChanged(backgroundColor => {
			domDraft.querySelector('#crater-employee-directory-search-type').css({ backgroundColor });
		});

		displayPane.querySelector('#Height-cell').onChanged(height => {
			domDraft.querySelector('.crater-employee-directory-display').css({ height });
		});

		displayPane.querySelector('#Background-Color-cell').onChanged(backgroundColor => {
			domDraft.querySelector('.crater-employee-directory-display').css({ backgroundColor });
		});

		displayPane.querySelector('#Font-Color-cell').onChanged(color => {
			domDraft.querySelector('.crater-employee-directory-display').css({ color });
		});

		displayPane.querySelector('#Font-Size-cell').onChanged(fontSize => {
			domDraft.querySelector('.crater-employee-directory-display').css({ fontSize });
		});

		displayPane.querySelector('#Font-Style-cell').onChanged(fontFamily => {
			domDraft.querySelector('.crater-employee-directory-display').css({ fontFamily });
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());
			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart

			this.sharePoint.properties.pane.content[this.key].settings.mailApp = appsPane.querySelector('#Mail-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.messageApp = appsPane.querySelector('#Message-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.callApp = appsPane.querySelector('#Call-cell').value;
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
		let controller = this.element.querySelector('#crater-carousel-controller'),
			arrows = this.element.querySelectorAll('.crater-arrow'),
			radios,
			columns = this.element.querySelectorAll('.crater-carousel-column'),
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
		radios = controller.querySelectorAll('.crater-carousel-radio-toggle');

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
			let columns = this.sharePoint.properties.pane.content[this.key].draft.dom.querySelectorAll('.crater-carousel-column');

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
						element: 'img', name: 'Image', attributes: {}, dataAttributes: { class: 'crater-icon', src: params.columns[i].querySelector('.crater-carousel-image').src }
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Text', attributes: {}, value: params.columns[i].querySelector('.crater-carousel-text').innerText || ''
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Color', attributes: {}, value: params.columns[i].querySelector('.crater-carousel-text').css().color || '', list: func.colors
					}),
					this.elementModifier.cell({
						element: 'input', name: 'BackgroundColor', attributes: {}, value: params.columns[i].querySelector('.crater-carousel-text').css().backgroundColor || ''
					}),
				]
			});
		}

		return columnsPane;
	}

	public listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		this.paneContent = this.sharePoint.app.querySelector('.crater-property-content').monitor();

		let domDraft = this.sharePoint.properties.pane.content[this.key].draft.dom;
		let content = domDraft.querySelector('.crater-carousel-content');
		let columns = domDraft.querySelectorAll('.crater-carousel-column');

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
				columnPane.querySelector('.crater-content-options').css({ visibility: 'visible' });
			});

			columnPane.addEventListener('mouseout', event => {
				columnPane.querySelector('.crater-content-options').css({ visibility: 'hidden' });
			});

			let imageCell = columnPane.querySelector('#Image-cell').parentNode;
			this.uploadImage({ parent: imageCell }, (image) => {
				imageCell.querySelector('#Image-cell').src = image.src;
				columnDom.querySelector('.crater-carousel-image').src = image.src;
			});

			columnPane.querySelector('#Text-cell').onChanged(value => {
				columnDom.querySelector('.crater-carousel-text').textContent = value;
			});

			let colorCell = columnPane.querySelector('#Color-cell').parentNode;
			this.pickColor({ parent: colorCell, cell: colorCell.querySelector('#Color-cell') }, (color) => {
				columnDom.querySelector('.crater-carousel-text').css({ color });
				colorCell.querySelector('#Color-cell').value = color;
				colorCell.querySelector('#Color-cell').setAttribute('value', color);
			});

			let backgroundColorCell = columnPane.querySelector('#BackgroundColor-cell').parentNode;
			this.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.querySelector('#BackgroundColor-cell') }, (backgroundColor) => {
				columnDom.css({ backgroundColor });
				backgroundColorCell.querySelector('#BackgroundColor-cell').value = backgroundColor;
				backgroundColorCell.querySelector('#BackgroundColor-cell').setAttribute('value', backgroundColor);
			});

			columnPane.querySelector('.delete-crater-carousel-column').addEventListener('click', event => {
				columnDom.remove();
				columnPane.remove();
			});

			columnPane.querySelector('.add-before-crater-carousel-column').addEventListener('click', event => {
				let newColumnPrototype = columnPrototype.cloneNode(true);
				let newColumnPanePrototype = columnPanePrototype.cloneNode(true);

				columnDom.before(newColumnPrototype);
				columnPane.before(newColumnPanePrototype);
				carouselColumnHandler(newColumnPanePrototype, newColumnPrototype);
			});

			columnPane.querySelector('.add-after-crater-carousel-column').addEventListener('click', event => {
				let newColumnPrototype = columnPrototype.cloneNode(true);
				let newColumnPanePrototype = columnPanePrototype.cloneNode(true);

				columnDom.after(newColumnPrototype);
				columnPane.after(newColumnPanePrototype);
				carouselColumnHandler(newColumnPanePrototype, newColumnPrototype);
			});
		};

		this.paneContent.querySelector('.new-component').addEventListener('click', event => {
			let newColumnPrototype = columnPrototype.cloneNode(true);
			let newColumnPanePrototype = columnPanePrototype.cloneNode(true);

			content.append(newColumnPrototype);//c
			this.paneContent.querySelector('.columns-pane').append(newColumnPanePrototype);

			carouselColumnHandler(newColumnPanePrototype, newColumnPrototype);
		});

		this.paneContent.querySelectorAll('.crater-carousel-column-pane').forEach((columnPane, position) => {
			carouselColumnHandler(columnPane, columns[position]);
		});

		let settingsPane = this.paneContent.querySelector('.settings-pane');

		settingsPane.querySelector('#Duration-cell').onChanged();

		settingsPane.querySelector('#Columns-cell').onChanged(value => {
			domDraft.querySelector('.crater-carousel-content').css({ gridTemplateColumns: `repeat(${value}, 1fr)` });
		});

		settingsPane.querySelector('#FontSize-cell').onChanged(fontSize => {
			domDraft.querySelectorAll('.crater-carousel-text').forEach(text => {
				text.css({ fontSize });
			});
		});

		settingsPane.querySelector('#FontStyle-cell').onChanged(fontFamily => {
			domDraft.querySelectorAll('.crater-carousel-text').forEach(text => {
				text.css({ fontFamily });
			});
		});

		settingsPane.querySelector('#ImageSize-cell').onChanged(width => {
			domDraft.querySelectorAll('.crater-carousel-image').forEach(text => {
				text.css({ width });
			});
		});

		settingsPane.querySelector('#ShowText-cell').onChanged(display => {
			domDraft.querySelectorAll('.crater-carousel-text').forEach(text => {
				if (display.toLowerCase() == 'no') {
					text.hide();
				} else {
					text.show();
				}
			});
		});

		settingsPane.querySelector('#Curved-cell').onChanged(curved => {
			domDraft.querySelectorAll('.crater-carousel-column').forEach(column => {
				if (curved.toLowerCase() == 'yes') {
					column.css({ borderRadius: '10px' });
				} else {
					column.cssRemove(['border-radius']);
				}
			});
		});

		settingsPane.querySelector('#Shadow-cell').onChanged(shadow => {
			domDraft.querySelectorAll('.crater-carousel-column').forEach(column => {
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

		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());
			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart

			this.sharePoint.properties.pane.content[this.key].settings.duration = this.paneContent.querySelector('#Duration-cell').value;
			this.sharePoint.properties.pane.content[this.key].settings.columns = this.paneContent.querySelector('#Columns-cell').value;
		});
	}

	public update(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;
		this.paneContent = this.setUpPaneContent(params);

		let paneConnection = this.sharePoint.app.querySelector('.crater-property-connection');

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
		updateWindow.querySelector('#update-element').addEventListener('click', event => {
			event.preventDefault();
			data.image = updateWindow.querySelector('#meta-data-image').value;
			data.text = updateWindow.querySelector('#meta-data-text').value;
			source = func.extractFromJsonArray(data, params.source);

			let newContent = this.render({ source });
			draftDom.querySelector('.crater-carousel-content').innerHTML = newContent.querySelector('.crater-carousel-content').innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = draftDom.outerHTML;

			this.paneContent.querySelector('.columns-pane').innerHTML = this.generatePaneContent({ columns: newContent.querySelectorAll('.crater-carousel-column') }).innerHTML;

			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
		});


		if (!func.isnull(paneConnection)) {
			paneConnection.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
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

		let content = events.querySelector('.crater-events-content');

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

		let rows = this.element.querySelectorAll('.crater-events-row');
		rows.forEach(row => {
			row.querySelector('.crater-events-row-icon').css({ gridColumnStart: iconPosition, gridColumnEnd: iconPosition, gridRowStart: 1 });
			row.querySelector('.crater-events-details').css({ gridColumnStart: detailsPosition, gridColumnEnd: detailsPosition, gridRowStart: 1 });
			row.querySelector('.crater-events-date').css({ gridColumnStart: datePosition, gridColumnEnd: datePosition, gridRowStart: 1 });
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
			let eventsContent = this.sharePoint.properties.pane.content[key].draft.dom.querySelector('.crater-events-content');
			let eventsRows = eventsContent.querySelectorAll('.crater-events-row');

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
								element: 'img', name: 'icon', attributes: {}, dataAttributes: { class: 'crater-icon', src: this.element.querySelector('.crater-events-title-icon').src }
							}),
							this.elementModifier.cell({
								element: 'input', name: 'iconsize', value: this.element.querySelector('.crater-events-title-icon').css()['width'] || ''
							}),
							this.elementModifier.cell({
								element: 'input', name: 'title', value: this.element.querySelector('.crater-events-title-text').textContent
							}),
							this.elementModifier.cell({
								element: 'input', name: 'backgroundcolor', value: this.element.querySelector('.crater-events-title').css()['background-color'], list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'color', value: this.element.querySelector('.crater-events-title').css().color, list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontsize', value: this.element.querySelector('.crater-events-title').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontstyle', value: this.element.querySelector('.crater-events-title').css()['font-family'] || ''
							}),
							this.elementModifier.cell({
								element: 'input', name: 'height', value: this.element.querySelector('.crater-events-title').css()['height'] || ''
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
						element: 'img', name: 'icon', attributes: {}, dataAttributes: { class: 'crater-icon', src: params.events[i].querySelector('#icon').src }
					}),
					this.elementModifier.cell({
						element: 'input', name: 'task', value: params.events[i].querySelector('#task').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'duration', value: params.events[i].querySelector('#duration').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'location', value: params.events[i].querySelector('#location').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'date', value: params.events[i].querySelector('#date').textContent
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
		this.paneContent = this.sharePoint.app.querySelector('.crater-property-content').monitor();

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

		let titlePane = this.paneContent.querySelector('.title-pane');

		titlePane.querySelector('#title-cell').onChanged(value => {
			domDraft.querySelector('.crater-events-title-text').textContent = value;
		});

		titlePane.querySelector('#fontsize-cell').onChanged(fontSize => {
			domDraft.querySelector('.crater-events-title-text').css({ fontSize });
		});

		titlePane.querySelector('#fontstyle-cell').onChanged(fontFamily => {
			domDraft.querySelector('.crater-events-title-text').css({ fontFamily });
		});

		let colorCell = titlePane.querySelector('#color-cell').parentNode;
		this.pickColor({ parent: colorCell, cell: colorCell.querySelector('#color-cell') }, (color) => {
			domDraft.querySelector('.crater-events-title-text').css({ color });
			colorCell.querySelector('#color-cell').value = color;
			colorCell.querySelector('#color-cell').setAttribute('value', color);
		});

		let iconCell = titlePane.querySelector('#icon-cell').parentNode;
		this.uploadImage({ parent: iconCell }, (image) => {
			iconCell.querySelector('#icon-cell').src = image.src;
			domDraft.querySelector('.crater-events-title-icon').src = image.src;
		});

		titlePane.querySelector('#iconsize-cell').onChanged(width => {
			domDraft.querySelector('.crater-events-title-icon').css({ width });
		});

		titlePane.querySelector('#height-cell').onChanged(height => {
			domDraft.querySelector('.crater-events-title').css({ height });
		});

		let backgroundColorCell = titlePane.querySelector('#backgroundcolor-cell').parentNode;
		this.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.querySelector('#backgroundcolor-cell') }, (backgroundColor) => {
			domDraft.querySelector('.crater-events-title').css({ backgroundColor });
			backgroundColorCell.querySelector('#backgroundcolor-cell').value = backgroundColor;
			backgroundColorCell.querySelector('#backgroundcolor-cell').setAttribute('value', backgroundColor);
		});

		let eventHandler = (eventPane, eventDom) => {
			eventPane.addEventListener('mouseover', event => {
				eventPane.querySelector('.crater-content-options').css({ visibility: 'visible' });
			});

			eventPane.addEventListener('mouseout', event => {
				eventPane.querySelector('.crater-content-options').css({ visibility: 'hidden' });
			});

			eventPane.querySelector('.delete-crater-event-row').addEventListener('click', event => {
				eventDom.remove();
				eventPane.remove();
			});

			eventPane.querySelector('.add-before-crater-event-row').addEventListener('click', event => {
				let newEventventPrototype = eventPrototype.cloneNode(true);
				let newEventPanePrototype = eventPanePrototype.cloneNode(true);

				eventDom.before(newEventventPrototype);
				eventPane.before(newEventPanePrototype);
				eventHandler(newEventPanePrototype, newEventventPrototype);
			});

			eventPane.querySelector('.add-after-crater-event-row').addEventListener('click', event => {
				let newEventventPrototype = eventPrototype.cloneNode(true);
				let newEventPanePrototype = eventPanePrototype.cloneNode(true);

				eventDom.after(newEventventPrototype);
				eventPane.after(newEventPanePrototype);
				eventHandler(newEventPanePrototype, newEventventPrototype);
			});

			let eventIconCell = eventPane.querySelector('#icon-cell').parentNode;
			this.uploadImage({ parent: eventIconCell }, (image) => {
				eventIconCell.querySelector('#icon-cell').src = image.src;
				domDraft.querySelector('.crater-events-row-icon').src = image.src;
			});

			eventPane.querySelector('#task-cell').onChanged(value => {
				eventDom.querySelector('.crater-events-task').innerText = value;
			});

			eventPane.querySelector('#duration-cell').onChanged(value => {
				eventDom.querySelector('.crater-events-duration').innerText = value;
			});

			eventPane.querySelector('#location-cell').onChanged(value => {
				eventDom.querySelector('.crater-events-location').innerText = value;
			});
			eventPane.querySelector('#date-cell').onChanged(value => {
				eventDom.querySelector('.crater-events-date').innerText = value;
			});
		};

		let eventsPane = this.paneContent.querySelector('.events-pane');

		this.paneContent.querySelector('.new-component').addEventListener('click', event => {
			let newevEntventPrototype = eventPrototype.cloneNode(true);
			let newEventPanePrototype = eventPanePrototype.cloneNode(true);

			eventsPane.append(newEventPanePrototype);
			domDraft.querySelector('.crater-events-content').append(newevEntventPrototype);
			eventHandler(newEventPanePrototype, newevEntventPrototype);
		});

		eventsPane.querySelectorAll('.crater-events-row-pane').forEach((eventPane, position) => {
			eventHandler(eventPane, domDraft.querySelectorAll('.crater-events-row')[position]);
		});

		let iconsPane = this.paneContent.querySelector('.icons-pane');

		iconsPane.querySelector('#size-cell').onChanged(width => {
			domDraft.querySelectorAll('.crater-events-row-icon').forEach(icon => {
				icon.css({ width });
			});
		});

		iconsPane.querySelector('#show-cell').onChanged(display => {
			domDraft.querySelectorAll('.crater-events-row-icon').forEach(icon => {
				if (display.toLowerCase() == 'no') icon.css({ display: 'none' });
				else icon.cssRemove(['display']);
			});
		});

		let handleProperties = (property, pane) => {
			pane.querySelector('#fontsize-cell').onChanged(fontSize => {
				domDraft.querySelectorAll(`.crater-events-${property}-value`).forEach(value => {
					value.css({ fontSize });
				});
			});

			pane.querySelector('#fontstyle-cell').onChanged(fontFamily => {
				domDraft.querySelectorAll(`.crater-events-${property}-value`).forEach(value => {
					value.css({ fontFamily });
				});
			});

			let paneColorCell = pane.querySelector('#color-cell').parentNode;
			this.pickColor({ parent: paneColorCell, cell: paneColorCell.querySelector('#color-cell') }, (color) => {
				domDraft.querySelectorAll(`.crater-events-${property}-value`).forEach(value => {
					value.css({ color });
				});
				paneColorCell.querySelector('#color-cell').value = color;
				paneColorCell.querySelector('#color-cell').setAttribute('value', color);
			});

			pane.querySelector('#show-cell').onChanged(display => {
				domDraft.querySelectorAll(`.crater-events-${property}`).forEach(aProperty => {
					if (display.toLowerCase() == 'no') aProperty.css({ display: 'none' });
					else aProperty.cssRemove(['display']);
				});
			});
		};

		let tasksPane = this.paneContent.querySelector('.tasks-pane');
		handleProperties('task', tasksPane);

		let durationsPane = this.paneContent.querySelector('.durations-pane');
		handleProperties('duration', durationsPane);

		let locationsPane = this.paneContent.querySelector('.locations-pane');
		handleProperties('location', locationsPane);

		let datesPane = this.paneContent.querySelector('.dates-pane');
		handleProperties('date', datesPane);

		let settingsPane = this.paneContent.querySelector('.settings-pane');

		settingsPane.querySelector('#Layout-cell').onChanged();

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		//on save clicked save the webpart settings and re-render
		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());
			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart			
			this.sharePoint.properties.pane.content[this.key].settings.layout = settingsPane.querySelector('#Layout-cell').value;
		});
	}

	public update(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;
		this.paneContent = this.setUpPaneContent(params);

		let paneConnection = this.sharePoint.app.querySelector('.crater-property-connection');

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
		updateWindow.querySelector('#update-element').addEventListener('click', event => {
			event.preventDefault();
			data.task = updateWindow.querySelector('#meta-data-task').value;
			data.icon = updateWindow.querySelector('#meta-data-icon').value;
			data.duration = updateWindow.querySelector('#meta-data-duration').value;
			data.location = updateWindow.querySelector('#meta-data-location').value;
			data.date = updateWindow.querySelector('#meta-data-date').value;
			source = func.extractFromJsonArray(data, params.source);

			let newContent = this.render({ source });
			draftDom.querySelector('.crater-events-content').innerHTML = newContent.querySelector('.crater-events-content').innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = draftDom.outerHTML;
			this.paneContent.querySelector('.events-pane').innerHTML = this.generatePaneContent({ events: newContent.querySelectorAll('.crater-events-row') }).innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
		});


		if (!func.isnull(paneConnection)) {
			paneConnection.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
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

		let buttons = this.element.querySelectorAll('.crater-button-single');

		let imageDisplay = this.sharePoint.properties.pane.content[this.key].settings.imageDisplay;
		let imageSize = this.sharePoint.properties.pane.content[this.key].settings.imageSize;
		let fontSize = this.sharePoint.properties.pane.content[this.key].settings.fontSize;
		let fontFamily = this.sharePoint.properties.pane.content[this.key].settings.fontFamily;
		let width = this.sharePoint.properties.pane.content[this.key].settings.width;
		let height = this.sharePoint.properties.pane.content[this.key].settings.height;

		buttons.forEach(button => {
			button.css({ height, width });
			button.querySelector('.crater-button-text').css({ fontSize, fontFamily });
			button.querySelector('.crater-button-icon').css({ width: imageSize });
			if (imageDisplay == 'No') button.querySelector('.crater-button-icon').hide();
			else button.querySelector('.crater-button-icon').show();
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
			let buttons = button.querySelectorAll('.crater-button-single');

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
						element: 'img', name: 'Image', attributes: {}, dataAttributes: { class: 'crater-icon', src: element.querySelector('.crater-button-icon').src }
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
		this.paneContent = this.sharePoint.app.querySelector('.crater-property-content').monitor();

		//fetch the content of Button
		let content = this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-button-content');
		let singleButtons = content.querySelectorAll('.crater-button-single');

		//fetch the icon
		let icons = content.querySelectorAll('.crater-button-icon');

		//fetch the text
		let texts = content.querySelectorAll('.crater-button-text');

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
				buttonPane.querySelector('.crater-content-options').css({ visibility: 'visible' });
			});

			buttonPane.addEventListener('mouseout', event => {
				buttonPane.querySelector('.crater-content-options').css({ visibility: 'hidden' });
			});

			buttonPane.querySelector('.delete-crater-single-button').addEventListener('click', event => {
				buttonDom.remove();
				buttonPane.remove();
			});

			buttonPane.querySelector('.add-before-crater-single-button').addEventListener('click', event => {
				let newButtonPrototype = buttonPrototype.cloneNode(true);
				let newButtonPanePrototype = buttonPanePrototype.cloneNode(true);

				buttonDom.before(newButtonPrototype);
				buttonPane.before(newButtonPanePrototype);
				buttonHandler(newButtonPanePrototype, newButtonPrototype);
			});

			buttonPane.querySelector('.add-after-crater-single-button').addEventListener('click', event => {
				let newButtonPrototype = buttonPrototype.cloneNode(true);
				let newButtonPanePrototype = buttonPanePrototype.cloneNode(true);

				buttonDom.after(newButtonPrototype);
				buttonPane.after(newButtonPanePrototype);
				buttonHandler(newButtonPanePrototype, newButtonPrototype);
			});

			buttonPane.querySelector('#Text-cell').onChanged(value => {
				buttonDom.querySelector('.crater-button-text').innerText = value;
			});

			buttonPane.querySelector('#Link-cell').onChanged(href => {
				buttonDom.setAttribute('href', href);
			});

			let colorCell = buttonPane.querySelector('#FontColor-cell').parentNode;
			this.pickColor({ parent: colorCell, cell: colorCell.querySelector('#FontColor-cell') }, (color) => {
				buttonDom.querySelector('.crater-button-text').css({ color }); colorCell.querySelector('#FontColor-cell').value = color;
				colorCell.querySelector('#FontColor-cell').setAttribute('value', color);
			});

			let backgroundColorCell = buttonPane.querySelector('#BackgroundColor-cell').parentNode;
			this.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.querySelector('#BackgroundColor-cell') }, (backgroundColor) => {
				buttonDom.css({ backgroundColor });
				backgroundColorCell.querySelector('#BackgroundColor-cell').value = backgroundColor;
				backgroundColorCell.querySelector('#BackgroundColor-cell').setAttribute('value', backgroundColor);
			});

			let imageCell = buttonPane.querySelector('#Image-cell').parentNode;
			//upload the icon
			this.uploadImage({ parent: imageCell }, (image) => {
				imageCell.querySelector('#Image-cell').src = image.src;
				buttonDom.querySelector('.crater-button-icon').src = image.src;
			});
		};

		this.paneContent.querySelector('.new-component').addEventListener('click', event => {
			let newButtonPrototype = buttonPrototype.cloneNode(true);
			let newButtonPanePrototype = buttonPanePrototype.cloneNode(true);

			content.append(newButtonPrototype);//c
			this.paneContent.querySelector('.content-pane').append(newButtonPanePrototype);

			buttonHandler(newButtonPanePrototype, newButtonPrototype);
		});

		this.paneContent.querySelectorAll('.button-pane').forEach((singlePane, position) => {
			buttonHandler(singlePane, singleButtons[position]);
		});

		let settingsPane = this.paneContent.querySelector('.settings-pane');

		settingsPane.querySelector('#Font-Size-cell').onChanged();

		settingsPane.querySelector('#Font-Family-cell').onChanged();

		settingsPane.querySelector('#Width-cell').onChanged();

		settingsPane.querySelector('#Height-cell').onChanged();

		//set the display of the button
		settingsPane.querySelector('#Image-Display-cell').onChanged();

		settingsPane.querySelector('#Image-Size-cell').onChanged();

		// on panecontent changed set the state
		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		//on save clicked save the webpart settings and re-render
		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());
			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart

			this.sharePoint.properties.pane.content[this.key].settings.ImageSize = settingsPane.querySelector('#Image-Size-cell').value;
			this.sharePoint.properties.pane.content[this.key].settings.imageDisplay = settingsPane.querySelector('#Image-Display-cell').value;
			this.sharePoint.properties.pane.content[this.key].settings.fontSize = settingsPane.querySelector('#Font-Size-cell').value;
			this.sharePoint.properties.pane.content[this.key].settings.fontFamily = settingsPane.querySelector('#Font-Family-cell').value;
			this.sharePoint.properties.pane.content[this.key].settings.width = settingsPane.querySelector('#Width-cell').value;
			this.sharePoint.properties.pane.content[this.key].settings.height = settingsPane.querySelector('#Height-cell').value;
		});
	}

	public update(params) {
		this.element = params.element;
		this.key = this.element.dataset.key;
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;
		this.paneContent = this.setUpPaneContent(params);

		let paneConnection = this.sharePoint.app.querySelector('.crater-property-connection');

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
		updateWindow.querySelector('#update-element').addEventListener('click', event => {
			event.preventDefault();
			data.image = updateWindow.querySelector('#meta-data-image').value;
			data.link = updateWindow.querySelector('#meta-data-link').value;
			data.title = updateWindow.querySelector('#meta-data-title').value;
			data.text = updateWindow.querySelector('#meta-data-text').value;
			source = func.extractFromJsonArray(data, params.source);

			let newContent = this.render({ source });
			draftDom.querySelector('.crater-button-content').innerHTML = newContent.querySelector('.crater-button-content').innerHTML;

			this.sharePoint.properties.pane.content[this.key].draft.html = draftDom.outerHTML;

			this.paneContent.querySelector('.content-pane').innerHTML = this.generatePaneContent({ buttons: newContent.querySelectorAll('.crater-button-single') }).innerHTML;

			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
		});

		if (!func.isnull(paneConnection)) {
			paneConnection.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {

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

		let icons = this.element.querySelectorAll('.crater-icons-icon');
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
						element: 'img', name: 'Image', attributes: {}, dataAttributes: { class: 'crater-icon', src: params.icons[i].querySelector('.crater-icons-icon-image').src }
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
			let icons = this.sharePoint.properties.pane.content[this.key].draft.dom.querySelectorAll('.crater-icons-icon');
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
		this.paneContent = this.sharePoint.app.querySelector('.crater-property-content').monitor();
		let content = this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-icons-content');
		let icons = this.sharePoint.properties.pane.content[this.key].draft.dom.querySelectorAll('.crater-icons-icon');

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
				iconPane.querySelector('.crater-content-options').css({ visibility: 'visible' });
			});

			iconPane.addEventListener('mouseout', event => {
				iconPane.querySelector('.crater-content-options').css({ visibility: 'hidden' });
			});

			let imageCell = iconPane.querySelector('#Image-cell').parentNode;
			this.uploadImage({ parent: imageCell }, (image) => {
				imageCell.querySelector('#Image-cell').src = image.src;
				iconDom.querySelector('.crater-icons-icon-image').src = image.src;
			});

			iconPane.querySelector('#Link-cell').onChanged(value => {
				iconDom.setAttribute('href', value);
			});

			iconPane.querySelector('#Title-cell').onChanged(value => {
				iconDom.setAttribute('title', value);
			});

			iconPane.querySelector('.delete-crater-icons-icon').addEventListener('click', event => {
				iconDom.remove();
				iconPane.remove();
			});

			iconPane.querySelector('.add-before-crater-icons-icon').addEventListener('click', event => {
				let newIconPrototype = iconPrototype.cloneNode(true);
				let newColumnPanePrototype = iconPanePrototype.cloneNode(true);

				iconDom.before(newIconPrototype);
				iconPane.before(newColumnPanePrototype);
				iconHandler(newColumnPanePrototype, newIconPrototype);
			});

			iconPane.querySelector('.add-after-crater-icons-icon').addEventListener('click', event => {
				let newIconPrototype = iconPrototype.cloneNode(true);
				let newColumnPanePrototype = iconPanePrototype.cloneNode(true);

				iconDom.after(newIconPrototype);
				iconPane.after(newColumnPanePrototype);
				iconHandler(newColumnPanePrototype, newIconPrototype);
			});
		};

		this.paneContent.querySelector('.new-component').addEventListener('click', event => {
			let newIconPrototype = iconPrototype.cloneNode(true);
			let newIconPanePrototype = iconPanePrototype.cloneNode(true);

			content.append(newIconPrototype);//c
			this.paneContent.querySelector('.counter-pane').append(newIconPanePrototype);

			iconHandler(newIconPanePrototype, newIconPrototype);
		});

		this.paneContent.querySelectorAll('.crater-icons-icon-pane').forEach((iconPane, position) => {
			iconHandler(iconPane, icons[position]);
		});

		this.paneContent.querySelector('#Width-cell').onChanged();
		this.paneContent.querySelector('#Height-cell').onChanged();
		this.paneContent.querySelector('#SpaceBetween-cell').onChanged();
		this.paneContent.querySelector('#Curved-cell').onChanged();


		let backgroundColorCell = this.paneContent.querySelector('#BackgroundColor-cell').parentNode;

		this.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.querySelector('#BackgroundColor-cell') }, (backgroundColor) => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.querySelectorAll('.crater-icons-icon').forEach(icon => {
				icon.css({
					backgroundColor
				});
			});
			backgroundColorCell.querySelector('#BackgroundColor-cell').value = backgroundColor;
			backgroundColorCell.querySelector('#BackgroundColor-cell').setAttribute('value', backgroundColor);
		});

		let colorCell = this.paneContent.querySelector('#Color-cell').parentNode;

		this.pickColor({ parent: colorCell, cell: colorCell.querySelector('#Color-cell') }, (color) => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.querySelectorAll('.crater-icons-icon').forEach(icon => {
				icon.css({
					color
				});
			});
			colorCell.querySelector('#Color-cell').value = color;
			colorCell.querySelector('#Color-cell').setAttribute('value', color);
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());
			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart

			this.element.querySelectorAll('.crater-icons-icon').forEach(icon => {
				if (this.paneContent.querySelector('#Curved-cell').value == 'Yes') {
					icon.classList.add('crater-curve');
				} else {
					icon.classList.remove('crater-curve');
				}

				icon.css({
					width: this.paneContent.querySelector('#Width-cell').value,
					height: this.paneContent.querySelector('#Height-cell').value,
					margin: this.paneContent.querySelector('#SpaceBetween-cell').value
				});
			});
		});
	}

	public update(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;
		this.paneContent = this.setUpPaneContent(params);

		let paneConnection = this.sharePoint.app.querySelector('.crater-property-connection');

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
		updateWindow.querySelector('#update-element').addEventListener('click', event => {
			event.preventDefault();
			data.image = updateWindow.querySelector('#meta-data-image').value;
			data.title = updateWindow.querySelector('#meta-data-title').value;
			data.link = updateWindow.querySelector('#meta-data-link').value;
			source = func.extractFromJsonArray(data, params.source);

			let newContent = this.render({ source });
			draftDom.querySelector('.crater-icons-content').innerHTML = newContent.querySelector('.crater-icons-content').innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = draftDom.outerHTML;

			this.paneContent.querySelector('.counter-pane').innerHTML = this.generatePaneContent({ icons: newContent.querySelectorAll('.crater-icons-icon') }).innerHTML;

			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
		});


		if (!func.isnull(paneConnection)) {
			paneConnection.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
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

			textAreaPane.innerHTML = this.element.querySelector('.crater-textarea-content').innerHTML;
		}

		return this.paneContent;
	}

	public listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset.key;
		this.paneContent = this.sharePoint.app.querySelector('.crater-property-content').monitor();

		let frBoxes = this.paneContent.querySelectorAll('.fr-box');

		for (let frBox of frBoxes) {
			frBox.remove();
		}

		this.paneContent.querySelector('textarea').innerHTML = this.element.querySelector('.crater-textarea-content').innerHTML;
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

		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart
			this.element.querySelector('.crater-textarea-content').innerHTML = '';
			let children = this.paneContent.querySelector('.fr-element').childNodes;
			for (let i of children) {
				this.element.querySelector('.crater-textarea-content').append(i);
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
		let menu = this.element.querySelector('.crater-section-menu');

		if (!func.isnull(menu) && func.isset(this.element.dataset.view) && this.element.dataset.view == 'Tabbed') {//if menu exists and section is tabbed
			menu.querySelectorAll('li').forEach(li => {
				let found = false;
				let owner = li.dataset.owner;
				for (let keyedElement of this.element.querySelector('.crater-section-content').querySelectorAll('.keyed-element')) {
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
					this.element.querySelectorAll('.keyed-element').forEach(keyedElement => {
						keyedElement.classList.add('in-active');
						if (li.dataset.owner == keyedElement.dataset.key) {
							keyedElement.classList.remove('in-active');
						}
					});
				}
			});
			if (this.element.dataset.view == 'Tabbed') {
				//click the last menu
				let menuButtons = menu.querySelectorAll('li');
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

			let elementContents = this.element.querySelector('.crater-section-content').querySelectorAll('.keyed-element');

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

		let contents = this.element.querySelector('.crater-section-content').querySelectorAll('.keyed-element');
		this.paneContent.querySelector('.section-contents-pane').innerHTML = this.generatePaneContent({ source: contents }).innerHTML;

		return this.paneContent;
	}

	private listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];

		this.paneContent = this.sharePoint.app.querySelector('.crater-property-content').monitor();

		let sectionContents = this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-section-content');

		let sectionContentDom = sectionContents.childNodes;
		let sectionContentPane = this.paneContent.querySelector('.section-contents-pane');

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
				sectionContentRowPane.querySelector('.crater-content-options').css({ visibility: 'visible' });
			});

			sectionContentRowPane.addEventListener('mouseout', event => {
				sectionContentRowPane.querySelector('.crater-content-options').css({ visibility: 'hidden' });
			});

			sectionContentRowPane.querySelector('.crater-section-webpart-name').textContent = sectionContentRowDom.dataset.type;

			sectionContentRowPane.querySelector('.delete-crater-section-content-row').addEventListener('click', event => {
				sectionContentRowDom.remove();
				sectionContentRowPane.remove();
			});

			sectionContentRowPane.querySelector('.add-before-crater-section-content-row').addEventListener('click', event => {
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

			sectionContentRowPane.querySelector('.add-after-crater-section-content-row').addEventListener('click', event => {
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
		this.paneContent.querySelector('.new-component').addEventListener('click', event => {
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

		this.paneContent.querySelectorAll('.crater-section-content-row-pane').forEach((sectionContent, position) => {
			//listen for events on all webparts
			sectionContentRowHandler(sectionContent, sectionContentDom[position]);
		});

		//monitor pane contents and note the changes
		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		//save the the noted changes when save button is clicked
		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;

			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());

			for (let keyedElement of this.element.querySelectorAll('.keyed-element')) {
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
		let menu = this.element.querySelector('.crater-menu');

		let list = [];
		let showIcons = this.sharePoint.properties.pane.content[this.key].settings.showMenuIcons;

		for (let keyedElement of this.element.querySelector('.crater-tab-content').childNodes) {
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

		menu.querySelectorAll('.crater-menu-item-icon').forEach(icon => {
			let width = this.sharePoint.properties.pane.content[this.key].settings.iconSize || '2em';
			icon.css({ width });
		});

		//onmneu clicked change to the webpart
		menu.addEventListener('click', event => {
			if (event.target.classList.contains('crater-menu-item')) {
				let item = event.target;
				for (let keyedElement of this.element.querySelector('.crater-tab-content').childNodes) {
					if (keyedElement.classList.contains('keyed-element')) {
						keyedElement.classList.add('in-active');
						if (item.dataset.owner == keyedElement.dataset.key) {
							keyedElement.classList.remove('in-active');
						}
					}
				}
			}
		});

		let menuButtons = menu.querySelectorAll('.crater-menu-item');
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

			let menus = tab.querySelector('.crater-menu');

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

			let elementContents = tab.querySelector('.crater-tab-content').childNodes;

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

		let contents = tab.querySelector('.crater-tab-content').childNodes;
		this.paneContent.querySelector('.tab-contents-pane').innerHTML = this.generatePaneContent({ source: contents }).innerHTML;

		return this.paneContent;
	}

	private listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];

		this.paneContent = this.sharePoint.app.querySelector('.crater-property-content').monitor();

		let menu = this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-menu');
		let tabContents = this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-tab-content');

		let tabContentDom = tabContents.childNodes;
		let tabContentPane = this.paneContent.querySelector('.tab-contents-pane');

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
				tabContentRowPane.querySelector('.crater-content-options').css({ visibility: 'visible' });
			});

			tabContentRowPane.addEventListener('mouseout', event => {
				tabContentRowPane.querySelector('.crater-content-options').css({ visibility: 'hidden' });
			});

			let iconCell = tabContentRowPane.querySelector('#Icon-cell').parentNode;

			this.uploadImage({ parent: iconCell }, (image) => {
				iconCell.querySelector('#Icon-cell').src = image.src;
				tabContentRowDom.dataset.icon = image.src;
			});

			tabContentRowPane.querySelector('#Title-cell').onChanged(value => {
				tabContentRowDom.dataset.title = value;
			});

			tabContentRowPane.querySelector('.delete-crater-tab-content-row').addEventListener('click', event => {
				tabContentRowDom.remove();
				tabContentRowPane.remove();
			});

			tabContentRowPane.querySelector('.add-before-crater-tab-content-row').addEventListener('click', event => {
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

			tabContentRowPane.querySelector('.add-after-crater-tab-content-row').addEventListener('click', event => {
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

		let menuPane = this.paneContent.querySelector('.menu-pane');

		menuPane.querySelector('#FontSize-cell').onChanged(fontSize => {
			menu.css({ fontSize });
		});

		menuPane.querySelector('#FontStyle-cell').onChanged(fontFamily => {
			menu.css({ fontFamily });
		});

		menuPane.querySelector('#IconSize-cell').onChanged();
		menuPane.querySelector('#ShowIcons-cell').onChanged();

		let backgroundColorCell = menuPane.querySelector('#BackgroundColor-cell').parentNode;
		this.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.querySelector('#BackgroundColor-cell') }, (backgroundColor) => {
			menu.css({ backgroundColor });
			backgroundColorCell.querySelector('#BackgroundColor-cell').value = backgroundColor;
			backgroundColorCell.querySelector('#BackgroundColor-cell').setAttribute('value', backgroundColor);
		});

		let colorCell = menuPane.querySelector('#Color-cell').parentNode;
		this.pickColor({ parent: colorCell, cell: colorCell.querySelector('#Color-cell') }, (color) => {
			menu.css({ color });
			colorCell.querySelector('#Color-cell').value = color;
			colorCell.querySelector('#Color-cell').setAttribute('value', color);
		});

		//add new webpart to the section
		this.paneContent.querySelector('.new-component').addEventListener('click', event => {
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

		this.paneContent.querySelectorAll('.crater-tab-content-row-pane').forEach((sectionContent, position) => {
			//listen for events on all webparts
			tabContentRowHandler(sectionContent, tabContentDom[position]);
		});

		//monitor pane contents and note the changes
		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		//save the the noted changes when save button is clicked
		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;

			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());

			for (let keyedElement of this.element.querySelectorAll('.keyed-element')) {
				this[keyedElement.dataset.type]({ action: 'rendered', element: keyedElement, sharePoint: this.sharePoint });
				console.log(keyedElement);

			}

			this.sharePoint.properties.pane.content[this.key].settings.showMenuIcons = menuPane.querySelector('#ShowIcons-cell').value;
			this.sharePoint.properties.pane.content[this.key].settings.iconSize = menuPane.querySelector('#IconSize-cell').value;
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
			this.element.querySelectorAll('.crater-arrow').forEach(arrow => {
				arrow.css({ visibility: 'visible' });
			});
			this.element.querySelector('.crater-top-right').css({ visibility: 'visible' });
		});

		//hide controllers and arrows
		this.element.addEventListener('mouseleave', () => {
			this.element.querySelectorAll('.crater-arrow').forEach(arrow => {
				arrow.css({ visibility: 'hidden' });
			});
			this.element.querySelector('.crater-top-right').css({ visibility: 'hidden' });
		});

		//make slides and images same as that of slider
		this.element.querySelectorAll('.crater-slide').forEach(slide => {
			slide.css({ height: this.element.position().height + 'px' });
			slide.querySelectorAll('img').forEach(img => {
				img.css({ height: this.element.position().height + 'px', filter });
			});
		});

		this.element.querySelectorAll('.crater-slide-quote').forEach(quote => {
			quote.css({ fontFamily: settings.textFontStyle, fontSize: settings.textFontSize, color: settings.textColor });
		});

		this.element.querySelectorAll('.crater-slide-link').forEach(link => {
			link.css({ fontFamily: settings.linkFontStyle, fontSize: settings.linkFontStyle, color: settings.linkColor, backgroundColor: settings.linkBackgroundColor, border: settings.linkBorder });

			if (settings.linkShow == 'No') {
				link.hide();
			} else {
				link.show();
			}
		});

		this.element.querySelectorAll('.crater-slide-sub-title').forEach(subTitle => {
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

		this.element.querySelectorAll('.crater-slide-details').forEach(detail => {
			detail.css({ alignSelf });
		});
	}

	//start the slider animation
	public startSlide() {
		this.key = this.element.dataset['key'];
		let controller = this.element.querySelector('#crater-slide-controller'),
			arrows = this.element.querySelectorAll('.crater-arrow'),
			radios,
			slides = this.element.querySelectorAll('.crater-slide'),
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
		radios = controller.querySelectorAll('.crater-slide-radio-toggle');

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

			let slides = this.sharePoint.properties.pane.content[key].draft.dom.querySelectorAll('.crater-slide');

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
						element: 'img', name: 'Image', attributes: {}, dataAttributes: { class: 'crater-icon', src: params.slides[i].querySelector('img').src }
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Quote', value: params.slides[i].querySelector('.crater-slide-quote').innerText
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Link', value: params.slides[i].querySelector('.crater-slide-link').href
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Link Text', value: params.slides[i].querySelector('.crater-slide-link').innerText
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Sub Title', value: params.slides[i].querySelector('.crater-slide-sub-title').innerText
					}),
				]
			});
		}
		return listPane;
	}

	private listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		this.paneContent = this.sharePoint.app.querySelector('.crater-property-content').monitor();

		let slides = this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-slides');

		let slideListRows = slides.querySelectorAll('.crater-slide');

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
				listRowPane.querySelector('.crater-content-options').css({ visibility: 'visible' });
			});

			listRowPane.addEventListener('mouseout', event => {
				listRowPane.querySelector('.crater-content-options').css({ visibility: 'hidden' });
			});

			let imageCell = listRowPane.querySelector('#Image-cell').parentNode;
			this.uploadImage({ parent: imageCell }, (image) => {
				imageCell.querySelector('#Image-cell').src = image.src;
				listRowDom.querySelector('img').src = image.src;
			});

			listRowPane.querySelector('#Quote-cell').onChanged(value => {
				listRowDom.querySelector('.crater-slide-quote').innerHTML = value;
			});

			listRowPane.querySelector('#Link-cell').onChanged(value => {
				listRowDom.querySelector('.crater-slide-link').href = value;
			});

			listRowPane.querySelector('#Link-Text-cell').onChanged(value => {
				listRowDom.querySelector('.crater-slide-link').innerText = value;
			});

			listRowPane.querySelector('#Sub-Title-cell').onChanged(value => {
				listRowDom.querySelector('.crater-slide-sub-title').innerText = value;
			});

			listRowPane.querySelector('.delete-crater-slide-content-row').addEventListener('click', event => {
				listRowDom.remove();
				listRowPane.remove();
			});

			listRowPane.querySelector('.add-before-crater-slide-content-row').addEventListener('click', event => {
				let newSlide = slidePrototype.cloneNode(true);
				let newListRow = listRowPrototype.cloneNode(true);

				listRowDom.before(newSlide);
				listRowPane.before(newListRow);
				listRowHandler(newListRow, newSlide);
			});

			listRowPane.querySelector('.add-after-crater-slide-content-row').addEventListener('click', event => {
				let newSlide = slidePrototype.cloneNode(true);
				let newListRow = listRowPrototype.cloneNode(true);

				listRowDom.after(newSlide);
				listRowPane.after(newListRow);

				listRowHandler(newListRow, newSlide);
			});
		};

		this.paneContent.querySelector('.new-component').addEventListener('click', event => {
			let newSlide = slidePrototype.cloneNode(true);
			let newListRow = listRowPrototype.cloneNode(true);

			slides.append(newSlide);
			this.paneContent.querySelector('.list-pane').append(newListRow);

			listRowHandler(newListRow, newSlide);
		});

		this.paneContent.querySelectorAll('.crater-slide-row-pane').forEach((listRow, position) => {
			listRowHandler(listRow, slideListRows[position]);
		});

		this.paneContent.querySelector('#Duration-cell').onChanged();
		this.paneContent.querySelector('#Content-Location-cell').onChanged();
		this.paneContent.querySelector('#View-cell').onChanged();
		this.paneContent.querySelector('#Image-Brightness-cell').onChanged();
		this.paneContent.querySelector('#Image-Blur-cell').onChanged();

		let textSettings = this.paneContent.querySelector('.text-settings');
		let linkSettings = this.paneContent.querySelector('.link-settings');
		let subTitleSettings = this.paneContent.querySelector('.sub-title-settings');

		textSettings.querySelector('#Font-Style-cell').onChanged();
		textSettings.querySelector('#Color-cell').onChanged();
		textSettings.querySelector('#Font-Size-cell').onChanged();

		linkSettings.querySelector('#Font-Style-cell').onChanged();
		linkSettings.querySelector('#Color-cell').onChanged();
		linkSettings.querySelector('#Font-Size-cell').onChanged();
		linkSettings.querySelector('#Show-cell').onChanged();
		linkSettings.querySelector('#Background-Color-cell').onChanged();
		linkSettings.querySelector('#Border-cell').onChanged();

		subTitleSettings.querySelector('#Font-Style-cell').onChanged();
		subTitleSettings.querySelector('#Color-cell').onChanged();
		subTitleSettings.querySelector('#Font-Size-cell').onChanged();
		subTitleSettings.querySelector('#Show-cell').onChanged();

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());

			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart

			this.sharePoint.properties.pane.content[this.key].settings.duration = this.paneContent.querySelector('#Duration-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.view = this.paneContent.querySelector('#View-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.imageBrightness = this.paneContent.querySelector('#Image-Brightness-cell').value + '%';

			this.sharePoint.properties.pane.content[this.key].settings.imageBlur = this.paneContent.querySelector('#Image-Blur-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.contentLocation = this.paneContent.querySelector('#Content-Location-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.textFontStyle = textSettings.querySelector('#Font-Style-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.textColor = textSettings.querySelector('#Color-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.textFontSize = textSettings.querySelector('#Font-Size-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.linkFontStyle = linkSettings.querySelector('#Font-Style-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.linkColor = linkSettings.querySelector('#Color-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.linkFontSize = linkSettings.querySelector('#Font-Size-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.linkShow = linkSettings.querySelector('#Show-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.linkBackgroundColor = linkSettings.querySelector('#Background-Color-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.linkBorder = linkSettings.querySelector('#Border-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.subTitleFontStyle = subTitleSettings.querySelector('#Font-Style-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.subTitleColor = subTitleSettings.querySelector('#Color-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.subTitleFontSize = subTitleSettings.querySelector('#Font-Size-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.subTitleShow = subTitleSettings.querySelector('#Show-cell').value;
		});
	}

	public update(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;
		this.paneContent = this.setUpPaneContent(params);

		let paneConnection = this.sharePoint.app.querySelector('.crater-property-connection');

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
		updateWindow.querySelector('#update-element').addEventListener('click', event => {
			event.preventDefault();
			data.image = updateWindow.querySelector('#meta-data-image').value;
			data.text = updateWindow.querySelector('#meta-data-text').value;

			source = func.extractFromJsonArray(data, params.source);

			let newContent = this.render({ source });
			draftDom.querySelector('.crater-slides').innerHTML = newContent.querySelector('.crater-slides').innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = draftDom.outerHTML;

			this.paneContent.querySelector('.list-pane').innerHTML = this.generatePaneContent({ slides: draftDom.querySelectorAll('.crater-slide') }).innerHTML;

			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
		});


		if (!func.isnull(paneConnection)) {
			paneConnection.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
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
			let peopleList = this.sharePoint.properties.pane.content[key].draft.dom.querySelector('.crater-list-content');
			let peopleListRows = peopleList.querySelectorAll('.crater-list-content-row');
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
								element: 'img', name: 'icon', attributes: {}, dataAttributes: { class: 'crater-icon', src: this.element.querySelector('.crater-list-title-icon').src }
							}),
							this.elementModifier.cell({
								element: 'input', name: 'title', value: this.element.querySelector('.crater-list-title').textContent
							}),
							this.elementModifier.cell({
								element: 'input', name: 'backgroundcolor', value: this.element.querySelector('.crater-list-title').css()['background-color'], list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'color', value: this.element.querySelector('.crater-list-title').css().color, list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontsize', value: this.element.querySelector('.crater-list-title').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'height', value: this.element.querySelector('.crater-list-title').css()['height'] || ''
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
						element: 'img', name: 'Image', attributes: {}, dataAttributes: { class: 'crater-icon', src: params.list[i].querySelector('#image').src }
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Title', value: params.list[i].querySelector('#title').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Job', value: params.list[i].querySelector('#job').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Link', value: params.list[i].querySelector('#link').href
					}),
				]
			});
		}

		return listPane;
	}

	public listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		this.paneContent = this.sharePoint.app.querySelector('.crater-property-content').monitor();

		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;

		let peopleList = draftDom.querySelector('.crater-list-content');
		let peopleListRows = peopleList.querySelectorAll('.crater-list-content-row');

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
				listRowPane.querySelector('.crater-content-options').css({ visibility: 'visible' });
			});

			listRowPane.addEventListener('mouseout', event => {
				listRowPane.querySelector('.crater-content-options').css({ visibility: 'hidden' });
			});

			let imageCell = listRowPane.querySelector('#Image-cell').parentNode;
			this.uploadImage({ parent: imageCell }, (image) => {
				imageCell.querySelector('#Image-cell').src = image.src;
				listRowDom.querySelector('.crater-list-content-row-image').src = image.src;
			});

			listRowPane.querySelector('#Title-cell').onChanged(value => {
				listRowDom.querySelector('.crater-list-content-row-details-title').innerHTML = value;
			});

			listRowPane.querySelector('#Job-cell').onChanged(value => {
				listRowDom.querySelector('.crater-list-content-row-details-job').innerHTML = value;
			});

			listRowPane.querySelector('#Link-cell').onChanged(value => {
				listRowDom.querySelector('.crater-list-content-row-details-link').href = value;
			});

			listRowPane.querySelector('.delete-crater-list-content-row').addEventListener('click', event => {
				listRowDom.remove();
				listRowPane.remove();
			});

			listRowPane.querySelector('.add-before-crater-list-content-row').addEventListener('click', event => {
				let newPeopleListRow = peopleContentRowPrototype.cloneNode(true);
				let newListRow = listRowPrototype.cloneNode(true);

				listRowDom.before(newPeopleListRow);
				listRowPane.before(newListRow);
				listRowHandler(newListRow, newPeopleListRow);
			});

			listRowPane.querySelector('.add-after-crater-list-content-row').addEventListener('click', event => {
				let newPeopleListRow = peopleContentRowPrototype.cloneNode(true);
				let newListRow = listRowPrototype.cloneNode(true);

				listRowDom.after(newPeopleListRow);
				listRowPane.after(newListRow);

				listRowHandler(newListRow, newPeopleListRow);
			});
		};

		let titlePane = this.paneContent.querySelector('.title-pane');
		let iconCell = titlePane.querySelector('#icon-cell').parentNode;

		this.uploadImage({ parent: iconCell }, (image) => {
			iconCell.querySelector('#icon-cell').src = image.src;
			this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-list-title-icon').src = image.src;
		});

		titlePane.querySelector('#title-cell').onChanged(value => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-list-title').innerHTML = value;
		});

		titlePane.querySelector('#Show-cell').onChanged(value => {
			if (value == 'No') {
				this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-list-title').hide();
			} else {
				this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-list-title').show();
			}
		});

		let backgroundColorCell = titlePane.querySelector('#backgroundcolor-cell').parentNode;
		this.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.querySelector('#backgroundcolor-cell') }, (backgroundColor) => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-list-title').css({ backgroundColor });
			backgroundColorCell.querySelector('#backgroundcolor-cell').value = backgroundColor;
			backgroundColorCell.querySelector('#backgroundcolor-cell').setAttribute('value', backgroundColor);
		});

		let colorCell = titlePane.querySelector('#color-cell').parentNode;
		this.pickColor({ parent: colorCell, cell: colorCell.querySelector('#color-cell') }, (color) => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-list-title').css({ color });
			colorCell.querySelector('#color-cell').value = color;
			colorCell.querySelector('#color-cell').setAttribute('value', color);
		});

		this.paneContent.querySelector('.title-pane').querySelector('#fontsize-cell').onChanged(value => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-list-title').css({ fontSize: value });
		});

		this.paneContent.querySelector('.title-pane').querySelector('#height-cell').onChanged(value => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-list-title').css({ height: value });
		});

		this.paneContent.querySelector('.new-component').addEventListener('click', event => {
			let newPeopleListRow = peopleContentRowPrototype.cloneNode(true);
			let newListRow = listRowPrototype.cloneNode(true);

			this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-list-content').append(newPeopleListRow);//c
			this.paneContent.querySelector('.list-pane').append(newListRow);

			listRowHandler(newListRow, newPeopleListRow);
		});

		this.paneContent.querySelectorAll('.crater-list-row-pane').forEach((listRow, position) => {
			listRowHandler(listRow, peopleListRows[position]);
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
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

		let paneConnection = this.sharePoint.app.querySelector('.crater-property-connection');

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
		updateWindow.querySelector('#update-element').addEventListener('click', event => {
			event.preventDefault();
			data.image = updateWindow.querySelector('#meta-data-image').value;
			data.title = updateWindow.querySelector('#meta-data-title').value;
			data.job = updateWindow.querySelector('#meta-data-job').value;
			data.link = updateWindow.querySelector('#meta-data-link').value;
			source = func.extractFromJsonArray(data, params.source);

			let newContent = this.render({ source });
			draftDom.querySelector('.crater-list-content').innerHTML = newContent.querySelector('.crater-list-content').innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = draftDom.outerHTML;

			this.paneContent.querySelector('.list-pane').innerHTML = this.generatePaneContent({ list: newContent.querySelectorAll('.crater-list-content-row') }).innerHTML;

			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
		});


		if (!func.isnull(paneConnection)) {
			paneConnection.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
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
		let tiles = this.element.querySelectorAll('.crater-tiles-content-column');
		let length: number = tiles.length;
		let currentContent;

		//fetch the settings
		this.columns = this.sharePoint.properties.pane.content[this.key].settings.columns / 1;
		this.duration = this.sharePoint.properties.pane.content[this.key].settings.duration;
		this.backgroundPosition = this.sharePoint.properties.pane.content[this.key].settings.backgroundPosition;
		this.backgroundWidth = this.sharePoint.properties.pane.content[this.key].settings.backgroundWidth;
		this.backgroundHeight = this.sharePoint.properties.pane.content[this.key].settings.backgroundHeight;

		this.height = this.element.css().height;

		this.element.querySelectorAll('.crater-tiles-content').forEach(content => {
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
			tile.querySelector('.crater-tiles-content-column-image').cssRemove(['margin-left']);
			tile.querySelector('.crater-tiles-content-column-image').cssRemove(['margin-right']);

			let getPosition = position => {
				if (position == 'left') return 'right';
				else if (position == 'right') return 'left';
				else return position;
			};

			let direction = getPosition(func.trem(this.backgroundPosition).toLowerCase());

			tileBackground[`margin-${direction}`] = 'auto';
			tileBackground.width = this.backgroundWidth;
			tileBackground.height = this.backgroundHeight;

			tile.querySelector('.crater-tiles-content-column-image').css(tileBackground);

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
			tile.querySelector('.crater-tiles-content-column-details').classList.add('crater-tiles-content-column-details-short');

			tile.querySelector('.crater-tiles-content-column-details').classList.remove('crater-tiles-content-column-details-full');

			//animate when hovered
			tile.addEventListener('mouseenter', event => {
				tile.querySelector('.crater-tiles-content-column-details').classList.add('crater-tiles-content-column-details-full');
				tile.querySelector('.crater-tiles-content-column-details').classList.remove('crater-tiles-content-column-details-short');
			});

			tile.addEventListener('mouseleave', event => {
				tile.querySelector('.crater-tiles-content-column-details').classList.add('crater-tiles-content-column-details-short');
				tile.querySelector('.crater-tiles-content-column-details').classList.remove('crater-tiles-content-column-details-full');
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
			let tiles = this.sharePoint.properties.pane.content[this.key].draft.dom.querySelectorAll('.crater-tiles-content-column');
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
		this.paneContent.querySelector('#Duration-cell').value = this.sharePoint.properties.pane.content[this.key].settings.duration || '';

		this.paneContent.querySelector('#Columns-cell').value = this.sharePoint.properties.pane.content[this.key].settings.columns || '';

		this.paneContent.querySelector('#BackgroundPosition-cell').value = this.sharePoint.properties.pane.content[this.key].settings.backgroundPosition || '';

		this.paneContent.querySelector('#BackgroundWidth-cell').value = this.sharePoint.properties.pane.content[this.key].settings.backgroundWidth || '';

		this.paneContent.querySelector('#BackgroundHeight-cell').value = this.sharePoint.properties.pane.content[this.key].settings.backgroundHeight || '';

		this.paneContent.querySelector('#Height-cell').value = this.sharePoint.properties.pane.content[this.key].settings.height || '';

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
						element: 'img', name: 'Image', attributes: {}, dataAttributes: { class: 'crater-icon', src: params.tiles[i].querySelector('.crater-tiles-content-column-image').src }
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Name', value: params.tiles[i].querySelector('#name').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'About', value: params.tiles[i].querySelector('#about').textContent
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

		this.paneContent = this.sharePoint.app.querySelector('.crater-property-content').monitor();

		let content = this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-tiles-content');
		let tiles = content.querySelectorAll('.crater-tiles-content-column');

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
				tilesColumnPane.querySelector('.crater-content-options').css({ visibility: 'visible' });
			});

			tilesColumnPane.addEventListener('mouseout', event => {
				tilesColumnPane.querySelector('.crater-content-options').css({ visibility: 'hidden' });
			});

			let imageCell = tilesColumnPane.querySelector('#Image-cell').parentNode;
			this.uploadImage({ parent: imageCell }, (image) => {
				imageCell.querySelector('#Image-cell').src = image.src;
				tilesColumnDom.querySelector('.crater-tiles-content-column-image').src = image.src;
			});


			tilesColumnPane.querySelector('#Name-cell').onChanged(value => {
				tilesColumnDom.querySelector('.crater-tiles-content-column-details-name').innerHTML = value;
			});

			tilesColumnPane.querySelector('#About-cell').onChanged(value => {
				tilesColumnDom.querySelector('.crater-tiles-content-column-details-about').innerHTML = value;
			});

			let colorCell = tilesColumnPane.querySelector('#Color-cell').parentNode;
			this.pickColor({ parent: colorCell, cell: colorCell.querySelector('#Color-cell') }, (color) => {
				tilesColumnDom.css({ color });
				colorCell.querySelector('#Color-cell').value = color;
				colorCell.querySelector('#Color-cell').setAttribute('value', color);
			});

			let backgroundColorCell = tilesColumnPane.querySelector('#Background-cell').parentNode;
			this.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.querySelector('#Background-cell') }, (backgroundColor) => {
				tilesColumnDom.css({ backgroundColor });
				backgroundColorCell.querySelector('#Background-cell').value = backgroundColor;
				backgroundColorCell.querySelector('#Background-cell').setAttribute('value', backgroundColor);
			});

			tilesColumnPane.querySelector('.delete-crater-tiles-content-column').addEventListener('click', event => {
				tilesColumnDom.remove();
				tilesColumnPane.remove();
			});

			tilesColumnPane.querySelector('.add-before-crater-tiles-content-column').addEventListener('click', event => {
				let newColumnPrototype = columnPrototype.cloneNode(true);
				let newColumnPanePrototype = columnPanePrototype.cloneNode(true);

				tilesColumnDom.before(newColumnPrototype);
				tilesColumnPane.before(newColumnPanePrototype);
				tilescolumnHandler(newColumnPanePrototype, newColumnPrototype);
			});

			tilesColumnPane.querySelector('.add-after-crater-tiles-content-column').addEventListener('click', event => {
				let newColumnPrototype = columnPrototype.cloneNode(true);
				let newColumnPanePrototype = columnPanePrototype.cloneNode(true);

				tilesColumnDom.after(newColumnPrototype);
				tilesColumnPane.after(newColumnPanePrototype);
				tilescolumnHandler(newColumnPanePrototype, newColumnPrototype);
			});
		};

		this.paneContent.querySelector('.new-component').addEventListener('click', event => {
			let newColumnPrototype = columnPrototype.cloneNode(true);
			let newColumnPanePrototype = columnPanePrototype.cloneNode(true);

			content.append(newColumnPrototype);//c
			this.paneContent.querySelector('.tiles-pane').append(newColumnPanePrototype);

			tilescolumnHandler(newColumnPanePrototype, newColumnPrototype);
		});

		this.paneContent.querySelectorAll('.crater-tiles-content-column-pane').forEach((tilesColumnPane, position) => {
			tilescolumnHandler(tilesColumnPane, tiles[position]);
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
			//update webpart            

			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());

			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;

			//save the new settings
			this.sharePoint.properties.pane.content[this.key].settings.duration = this.paneContent.querySelector('#Duration-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.columns = this.paneContent.querySelector('#Columns-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.backgroundPosition = this.paneContent.querySelector('#BackgroundPosition-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.backgroundHeight = this.paneContent.querySelector('#BackgroundHeight-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.backgroundWidth = this.paneContent.querySelector('#BackgroundWidth-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.height = this.paneContent.querySelector('#Height-cell').value;
		});
	}

	public update(params) {
		this.element = params.element;
		this.key = this.element.dataset.key;
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;
		this.paneContent = this.setUpPaneContent(params);

		let paneConnection = this.sharePoint.app.querySelector('.crater-property-connection');

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
		updateWindow.querySelector('#update-element').addEventListener('click', event => {
			event.preventDefault();
			data.image = updateWindow.querySelector('#meta-data-image').value;
			data.name = updateWindow.querySelector('#meta-data-name').value;
			data.about = updateWindow.querySelector('#meta-data-about').value;
			data.color = updateWindow.querySelector('#meta-data-color').value;
			source = func.extractFromJsonArray(data, params.source);

			let newContent = this.render({ source });
			draftDom.querySelector('.crater-tiles-content').innerHTML = newContent.querySelector('.crater-tiles-content').innerHTML;

			this.sharePoint.properties.pane.content[this.key].draft.html = draftDom.outerHTML;

			this.paneContent.querySelector('.tiles-pane').innerHTML = this.generatePaneContent({ tiles: newContent.querySelectorAll('.crater-tiles-content-column') }).innerHTML;

			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
		});

		if (!func.isnull(paneConnection)) {
			paneConnection.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {

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

		let counters = this.element.querySelectorAll('.crater-counter-content-column');
		let length = counters.length;
		let currentContent;

		this.columns = this.sharePoint.properties.pane.content[this.key].settings.columns / 1;
		this.duration = this.sharePoint.properties.pane.content[this.key].settings.duration / 1;
		this.backgroundPosition = this.sharePoint.properties.pane.content[this.key].settings.backgroundPosition;
		this.backgroundWidth = this.sharePoint.properties.pane.content[this.key].settings.backgroundWidth;
		this.backgroundHeight = this.sharePoint.properties.pane.content[this.key].settings.backgroundHeight;

		this.height = this.element.css().height;

		this.element.querySelectorAll('.crater-counter-content').forEach(content => {
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
			counter.querySelector('.crater-counter-content-column-image').css({ height: this.backgroundHeight, width: this.backgroundWidth, filter: `blur(${settings.backgroundFilter})` });

			if (settings.showIcons == 'No') {
				counter.querySelector('.crater-counter-content-column-image').hide();
			} else {
				counter.querySelector('.crater-counter-content-column-image').show();
			}

			if (this.backgroundPosition != 'Right') {
				counter.querySelector('.crater-counter-content-column-image').css({ gridColumnStart: 1, gridRowStart: 1 });
				counter.querySelector('.crater-counter-content-column-details').css({ gridColumnStart: 2, gridRowStart: 1 });

			} else {
				counter.querySelector('.crater-counter-content-column-image').css({ gridColumnStart: 2, gridRowStart: 1 });
				counter.querySelector('.crater-counter-content-column-details').css({ gridColumnStart: 1, gridRowStart: 1 });
			}

			let count = counter.querySelector('#count').dataset['count'];

			counter.querySelector('#count').innerHTML = 0;
			counter.querySelector('#unit').css({ visibility: 'hidden' });

			let interval = setInterval(() => {
				counter.querySelector('#count').innerHTML = counter.querySelector('#count').innerHTML / 1 + 1 / 1;
				if (counter.querySelector('#count').innerHTML == count) {
					counter.querySelector('#unit').css({ visibility: 'unset' });
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
			let counters = this.sharePoint.properties.pane.content[this.key].draft.dom.querySelectorAll('.crater-counter-content-column');

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

		this.paneContent.querySelector('#Box-Height-cell').value = this.sharePoint.properties.pane.content[this.key].settings.height || '';

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
						element: 'img', name: 'Image', attributes: {}, dataAttributes: { class: 'crater-icon', src: params.counters[i].querySelector('.crater-counter-content-column-image').src }
					}),
					this.elementModifier.cell({
						element: 'img', name: 'BackgroundImage', attributes: {}, dataAttributes: { class: 'crater-icon', src: params.counters[i].css()['background-image'] || '' }
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Count', value: params.counters[i].querySelector('#count').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Unit', value: params.counters[i].querySelector('#unit').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Name', value: params.counters[i].querySelector('#name').textContent
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
		this.paneContent = this.sharePoint.app.querySelector('.crater-property-content').monitor();
		let content = this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-counter-content');
		let counters = this.sharePoint.properties.pane.content[this.key].draft.dom.querySelectorAll('.crater-counter-content-column');

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
				counterColumnPane.querySelector('.crater-content-options').css({ visibility: 'visible' });
			});

			counterColumnPane.addEventListener('mouseout', event => {
				counterColumnPane.querySelector('.crater-content-options').css({ visibility: 'hidden' });
			});

			counterColumnPane.querySelector('#Image-cell').onChanged(value => {
				counterColumnDom.css({ backgroundImage: `url('${value}')` });
			});

			let imageCell = counterColumnPane.querySelector('#Image-cell').parentNode;
			this.uploadImage({ parent: imageCell }, (image) => {
				imageCell.querySelector('#Image-cell').src = image.src;
				counterColumnDom.querySelector('.crater-counter-content-column-image').src = image.src;
			});

			let backgroundImageCell = counterColumnPane.querySelector('#BackgroundImage-cell').parentNode;
			this.uploadImage({ parent: backgroundImageCell }, (backgroundImage) => {
				backgroundImageCell.querySelector('#BackgroundImage-cell').src = backgroundImage.src;
				counterColumnDom.setBackgroundImage(backgroundImage.src);
			});

			counterColumnPane.querySelector('#Count-cell').onChanged(value => {
				counterColumnDom.querySelector('.crater-counter-content-column-details-value-count').dataset['count'] = value;
			});

			counterColumnPane.querySelector('#Unit-cell').onChanged(value => {
				counterColumnDom.querySelector('.crater-counter-content-column-details-value-unit').innerHTML = value;
			});

			counterColumnPane.querySelector('#Name-cell').onChanged(value => {
				counterColumnDom.querySelector('.crater-counter-content-column-details-name').innerHTML = value;
			});

			let colorCell = counterColumnPane.querySelector('#Color-cell').parentNode;
			this.pickColor({ parent: colorCell, cell: colorCell.querySelector('#Color-cell') }, (color) => {
				counterColumnDom.css({ color });
				colorCell.querySelector('#Color-cell').value = color;
				colorCell.querySelector('#Color-cell').setAttribute('value', color);
			});

			let backgroundColorCell = counterColumnPane.querySelector('#Background-cell').parentNode;
			this.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.querySelector('#Background-cell') }, (backgroundColor) => {
				counterColumnDom.css({ backgroundColor });
				backgroundColorCell.querySelector('#Background-cell').value = backgroundColor;
				backgroundColorCell.querySelector('#Background-cell').setAttribute('value', backgroundColor);
			});

			counterColumnPane.querySelector('.delete-crater-counter-content-column').addEventListener('click', event => {
				counterColumnDom.remove();
				counterColumnPane.remove();
			});

			counterColumnPane.querySelector('.add-before-crater-counter-content-column').addEventListener('click', event => {
				let newColumnPrototype = columnPrototype.cloneNode(true);
				let newColumnPanePrototype = columnPanePrototype.cloneNode(true);

				counterColumnDom.before(newColumnPrototype);
				counterColumnPane.before(newColumnPanePrototype);
				countercolumnHandler(newColumnPanePrototype, newColumnPrototype);
			});

			counterColumnPane.querySelector('.add-after-crater-counter-content-column').addEventListener('click', event => {
				let newColumnPrototype = columnPrototype.cloneNode(true);
				let newColumnPanePrototype = columnPanePrototype.cloneNode(true);

				counterColumnDom.after(newColumnPrototype);
				counterColumnPane.after(newColumnPanePrototype);
				countercolumnHandler(newColumnPanePrototype, newColumnPrototype);
			});
		};

		this.paneContent.querySelector('.new-component').addEventListener('click', event => {
			let newColumnPrototype = columnPrototype.cloneNode(true);
			let newColumnPanePrototype = columnPanePrototype.cloneNode(true);

			content.append(newColumnPrototype);//c
			this.paneContent.querySelector('.counter-pane').append(newColumnPanePrototype);

			countercolumnHandler(newColumnPanePrototype, newColumnPrototype);
		});

		this.paneContent.querySelectorAll('.crater-counter-content-column-pane').forEach((counterColumnPane, position) => {
			countercolumnHandler(counterColumnPane, counters[position]);
		});

		this.paneContent.querySelector('#Duration-cell').onChanged();
		this.paneContent.querySelector('#Columns-cell').onChanged();
		this.paneContent.querySelector('#Box-Height-cell').onChanged();
		this.paneContent.querySelector('#Gap-cell').onChanged();
		this.paneContent.querySelector('#Show-Icons-cell').onChanged();
		this.paneContent.querySelector('#Background-Filter-cell').onChanged();
		this.paneContent.querySelector('#BackgroundWidth-cell').onChanged();
		this.paneContent.querySelector('#BackgroundHeight-cell').onChanged();

		this.sharePoint.properties.pane.content[this.key].settings.backgroundPosition = this.paneContent.querySelector('#BackgroundPosition-cell').value;

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());
			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart

			this.sharePoint.properties.pane.content[this.key].settings.duration = this.paneContent.querySelector('#Duration-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.columns = this.paneContent.querySelector('#Columns-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.height = this.paneContent.querySelector('#Box-Height-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.gap = this.paneContent.querySelector('#Gap-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.showIcons = this.paneContent.querySelector('#Show-Icons-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.backgroundFilter = this.paneContent.querySelector('#Background-Filter-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.backgroundWidth = this.paneContent.querySelector('#BackgroundWidth-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.backgroundHeight = this.paneContent.querySelector('#BackgroundHeight-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.backgroundPosition = this.paneContent.querySelector('#BackgroundPosition-cell').value;

			this.element.querySelectorAll('.crater-counter-content-column').forEach((element, position) => {
				let pane = this.paneContent.querySelectorAll('.crater-counter-content-column-pane')[position];

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

		let paneConnection = this.sharePoint.app.querySelector('.crater-property-connection');

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
		updateWindow.querySelector('#update-element').addEventListener('click', event => {
			event.preventDefault();
			data.image = updateWindow.querySelector('#meta-data-image').value;
			data.name = updateWindow.querySelector('#meta-data-name').value;
			data.count = updateWindow.querySelector('#meta-data-count').value;
			data.unit = updateWindow.querySelector('#meta-data-unit').value;
			data.color = updateWindow.querySelector('#meta-data-color').value;
			source = func.extractFromJsonArray(data, params.source);

			let newContent = this.render({ source });
			draftDom.querySelector('.crater-counter-content').innerHTML = newContent.querySelector('.crater-counter-content').innerHTML;

			this.sharePoint.properties.pane.content[this.key].draft.html = draftDom.outerHTML;

			this.paneContent.querySelector('.counter-pane').innerHTML = this.generatePaneContent({ counters: newContent.querySelectorAll('.crater-counter-content-column') }).innerHTML;

			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
		});

		if (!func.isnull(paneConnection)) {
			paneConnection.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {

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

		let newsContainer = news.querySelector('.crater-ticker-news-container');

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
		this.element.querySelector('.crater-ticker-title').css({ height: this.element.position().height + 'px' });
	}

	public startSlide() {
		this.key = this.element.dataset['key'];

		let news = this.element.querySelectorAll('.crater-ticker-news'),
			action = this.sharePoint.properties.pane.content[this.key].settings.animationType.toLowerCase();

		if (news.length < 2) return;

		let current = 0,
			key = 0,
			title = this.element.querySelector('.crater-ticker-title-text').position();

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

		this.element.querySelectorAll('.crater-arrow').forEach(arrow => {
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

		this.element.querySelectorAll('.crater-arrow')[current].click();

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
								element: 'input', name: 'Title', value: this.element.querySelector('.crater-ticker-title-text').innerText
							}),
							this.elementModifier.cell({
								element: 'input', name: 'BackgroundColor', value: this.element.querySelector('.crater-ticker-title').css()['background-color'], list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'TextColor', value: this.element.querySelector('.crater-ticker-title').css()['color'], list: func.colors
							}),
						]
					}
				]
			});

			let news = this.sharePoint.properties.pane.content[this.key].draft.dom.querySelectorAll('.crater-ticker-news');

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
		this.paneContent = this.sharePoint.app.querySelector('.crater-property-content').monitor();
		let domDraft = this.sharePoint.properties.pane.content[this.key].draft.dom;

		let tickerNewsContainer = domDraft.querySelector('.crater-ticker-news-container');

		let news = tickerNewsContainer.querySelectorAll('.crater-ticker-news');

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
				newsPane.querySelector('.crater-content-options').css({ visibility: 'visible' });
			});

			newsPane.addEventListener('mouseout', event => {
				newsPane.querySelector('.crater-content-options').css({ visibility: 'hidden' });
			});

			newsPane.querySelector('#Link-cell').onChanged(value => {
				newsDom.setAttribute('href', value);
			});

			newsPane.querySelector('#Text-cell').onChanged(value => {
				newsDom.dataset.text = value;
			});


			newsPane.querySelector('.delete-crater-ticker-content-row').addEventListener('click', event => {
				newsDom.remove();
				newsPane.remove();
			});

			newsPane.querySelector('.add-before-crater-ticker-content-row').addEventListener('click', event => {
				let newSlide = newsPrototype.cloneNode(true);
				let newListRow = newsPanePrototye.cloneNode(true);

				newsDom.before(newSlide);
				newsPane.before(newListRow);
				newsHandler(newListRow, newSlide);
			});

			newsPane.querySelector('.add-after-crater-ticker-content-row').addEventListener('click', event => {
				let newSlide = newsPrototype.cloneNode(true);
				let newListRow = newsPanePrototye.cloneNode(true);

				newsDom.after(newSlide);
				newsPane.after(newListRow);

				newsHandler(newListRow, newSlide);
			});
		};

		this.paneContent.querySelector('#Animation-cell').onChanged();
		this.paneContent.querySelector('#Duration-cell').onChanged();
		this.paneContent.querySelector('#View-cell').onChanged();

		this.paneContent.querySelector('.new-component').addEventListener('click', event => {
			let newSlide = newsPrototype.cloneNode(true);
			let newListRow = newsPanePrototye.cloneNode(true);

			tickerNewsContainer.append(newSlide);//c
			this.paneContent.querySelector('.news-pane').append(newListRow);

			newsHandler(newListRow, newSlide);
		});

		let backgroundColorCell = this.paneContent.querySelector('.title-pane').querySelector('#BackgroundColor-cell').parentNode;
		this.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.querySelector('#BackgroundColor-cell') }, (backgroundColor) => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-ticker-title').css({ backgroundColor });
			backgroundColorCell.querySelector('#BackgroundColor-cell').value = backgroundColor;
			backgroundColorCell.querySelector('#BackgroundColor-cell').setAttribute('value', backgroundColor);
		});

		let colorCell = this.paneContent.querySelector('.title-pane').querySelector('#TextColor-cell').parentNode;
		this.pickColor({ parent: colorCell, cell: colorCell.querySelector('#TextColor-cell') }, (color) => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-ticker-title').css({ color });
			colorCell.querySelector('#TextColor-cell').value = color;
			colorCell.querySelector('#TextColor-cell').setAttribute('value', color);
		});

		this.paneContent.querySelector('.title-pane').querySelector('#Title-cell').onChanged(value => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-ticker-title-text').innerText = value;
		});

		this.paneContent.querySelectorAll('.crater-ticker-news-pane').forEach((newsPane, position) => {
			newsHandler(newsPane, news[position]);
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());

			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart

			this.sharePoint.properties.pane.content[this.key].settings.duration = this.paneContent.querySelector('#Duration-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.animationType = this.paneContent.querySelector('#Animation-cell').value;

			this.sharePoint.properties.pane.content[this.key].settings.view = this.paneContent.querySelector('#View-cell').value;
		});
	}

	public update(params) {
		this.element = params.element;
		this.key = this.element.dataset.key;
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;
		this.paneContent = this.setUpPaneContent(params);

		let paneConnection = this.sharePoint.app.querySelector('.crater-property-connection');

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
		updateWindow.querySelector('#update-element').addEventListener('click', event => {
			event.preventDefault();
			data.link = updateWindow.querySelector('#meta-data-link').value;
			data.details = updateWindow.querySelector('#meta-data-details').value;
			source = func.extractFromJsonArray(data, params.source);

			let newContent = this.render({ source });
			draftDom.querySelector('.crater-ticker-news-container').innerHTML = newContent.querySelector('.crater-ticker-news-container').innerHTML;

			this.sharePoint.properties.pane.content[this.key].draft.html = draftDom.outerHTML;

			this.paneContent.querySelector('.news-pane').innerHTML = this.generatePaneContent({ news: newContent.querySelectorAll('.crater-ticker-news') }).innerHTML;

			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
		});

		if (!func.isnull(paneConnection)) {
			paneConnection.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {

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
			element: 'div', attributes: { class: 'crater-crater crater-component', style: { display: 'block', minHeight: '100px', width: '100%' }, 'data-type': 'crater' }, options: ['Edit', 'Delete'], children: [
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
		let sections = this.element.querySelectorAll('section.crater-section');
		let currentSection: any;
		let currentSibling: any;
		let otherDirection: any;
		let craterSectionsContainer = this.element.querySelector('.crater-sections-container');

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
			this.element.querySelectorAll('.crater-section').forEach(section => {
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
			this.element.querySelectorAll('.crater-section').forEach(section => {
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
			let container = this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-sections-container');

			let elementContents = container.querySelectorAll('.crater-section');

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
		this.paneContent.querySelector('#Columns-cell').value = this.sharePoint.properties.pane.content[this.key].settings.columns;

		this.paneContent.querySelector('#Columns-Sizes-cell').value = this.sharePoint.properties.pane.content[this.key].settings.columnsSizes;

		let contents = this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-sections-container').querySelectorAll('.crater-section');

		this.paneContent.querySelector('.sections-pane').innerHTML = this.generatePaneContent({ source: contents }).innerHTML;

		return this.paneContent;
	}

	public listenPaneContent(params?) {
		this.key = params.element.dataset['key'];
		this.element = params.element;
		this.paneContent = this.sharePoint.app.querySelector('.crater-property-content');
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;

		let sectionRowPanes = this.paneContent.querySelectorAll('.crater-section-row-pane');
		let sections = draftDom.querySelectorAll('.crater-section');

		let settingsPane = this.paneContent.querySelector('.settings-pane');

		settingsPane.querySelector('#Columns-Sizes-cell').onChanged();

		settingsPane.querySelector('#Columns-cell').onChanged(value => {
			settingsPane.querySelector('#Columns-Sizes-cell').setAttribute('value', `repeat(${value}, 1fr)`);
			settingsPane.querySelector('#Columns-Sizes-cell').value = `repeat(${value}, 1fr)`;
		});

		settingsPane.querySelector('#Scroll-cell').onChanged(scroll => {
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

		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());

			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart

			this.sharePoint.properties.pane.content[this.key].settings.columnsSizes = settingsPane.querySelector('#Columns-Sizes-cell').value;

			if (this.sharePoint.properties.pane.content[this.key].settings.columns < this.paneContent.querySelector('#Columns-cell').value) {
				this.sharePoint.properties.pane.content[this.key].settings.columns = this.paneContent.querySelector('#Columns-cell').value;
				this.sharePoint.properties.pane.content[this.key].settings.resetWidth = true;

				this.resetSections({ resetWidth: true });
			}
			else if (this.sharePoint.properties.pane.content[this.key].settings.columns > this.paneContent.querySelector('#Columns-cell').value) { //check if the columns is less than current
				alert("New number of column should be more than current");
			}
		});
	}

	private resetSections(params) {
		params = func.isset(params) ? params : {};
		let craterSectionsContainer = this.element.querySelector('.crater-sections-container');
		let sections = craterSectionsContainer.querySelectorAll('.crater-section');
		let count = sections.length;

		let number = this.sharePoint.properties.pane.content[this.key].settings.columns - count;
		let newSections = this.createSections({ number, height: '100px' }).querySelectorAll('.crater-section');
		//copy the current contents of the sections into the newly created sections
		for (let i = 0; i < newSections.length; i++) {
			craterSectionsContainer.append(newSections[i]);
		}

		// reset count
		count = craterSectionsContainer.querySelectorAll('.crater-section').length;

		craterSectionsContainer.css({ gridTemplateColumns: `repeat(${count}, 1fr` });
		craterSectionsContainer.querySelectorAll('.crater-section').forEach((section, position) => {
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

		let headers = table.querySelectorAll('th');
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
			let table = this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('table');

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

		this.paneContent.querySelector('tbody').querySelectorAll('tr').forEach(tr => {
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
			tablePane.querySelector('thead').innerHTML = params.header;
		}

		tablePane.querySelector('thead').querySelectorAll('th').forEach(th => {
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

		tablePane.querySelector('tbody').querySelectorAll('tr').forEach(tr => {
			tr.querySelectorAll('td').forEach(td => {
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
		this.paneContent = this.sharePoint.app.querySelector('.crater-property-content').monitor();
		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;
		let table = draftDom.querySelector('table');
		let tableBody = table.querySelector('tbody');

		let tableRows = tableBody.querySelectorAll('tr');

		let tableRowHandler = (tableRowPane, tableRowDom) => {
			tableRowPane.addEventListener('mouseover', event => {
				tableRowPane.querySelector('.crater-content-options').css({ visibility: 'visible' });
			});

			tableRowPane.addEventListener('mouseout', event => {
				tableRowPane.querySelector('.crater-content-options').css({ visibility: 'hidden' });
			});

			tableRowPane.querySelector('.delete-crater-table-content-row').addEventListener('click', event => {
				tableRowDom.remove();
				tableRowPane.remove();
			});

			tableRowPane.querySelector('.add-before-crater-table-content-row').addEventListener('click', event => {
				let newRow = tableRowDom.cloneNode(true);
				let newRowPane = tableRowPane.cloneNode(true);

				tableRowDom.before(newRow);
				tableRowPane.before(newRowPane);
				tableRowHandler(newRowPane, newRow);
			});

			tableRowPane.querySelector('.add-after-crater-table-content-row').addEventListener('click', event => {
				let newRow = tableRowDom.cloneNode(true);
				let newRowPane = tableRowPane.cloneNode(true);

				tableRowDom.after(newRow);
				tableRowPane.after(newRowPane);
				tableRowHandler(newRowPane, newRow);
			});

			tableRowPane.querySelectorAll('td').forEach((td, position) => {
				td.querySelector('input').onChanged(value => {
					tableRowDom.querySelectorAll('td')[position].textContent = value;
				});
			});
		};

		let dataName = 'crater-table-data-sample';//sample name

		this.paneContent.querySelector('tbody').querySelectorAll('tr').forEach((tableRow, position) => {
			tableRowHandler(tableRow, tableRows[position]);
		});

		let tableHeadHandler = (thPane, thDom) => {
			thPane.addEventListener('mouseover', event => {
				thPane.querySelector('.crater-content-options').css({ visibility: 'visible' });
				thPane.querySelector('.crater-table-sorter').css({ visibility: 'visible' });
			});

			thPane.addEventListener('mouseout', event => {
				thPane.querySelector('.crater-content-options').css({ visibility: 'hidden' });
				thPane.querySelector('.crater-table-sorter').css({ visibility: 'hidden' });
			});

			thPane.querySelector('.crater-table-sorter').addEventListener('click', event => {
				let order = thPane.querySelector('.crater-table-sorter').classList.contains('crater-up-arrow') ? -1 : 1;
				let name = thPane.dataset.name.split('crater-table-data-')[1];
				let data = this.elementModifier.sortTable(table, name, order);
				let newContent = this.render({ source: data });

				this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-table').innerHTML = newContent.querySelector('.crater-table').innerHTML;

				thPane.querySelector('.crater-table-sorter').classList.toggle('crater-up-arrow');
				thPane.querySelector('.crater-table-sorter').classList.toggle('crater-down-arrow');

				this.sharePoint.properties.pane.content[this.key].settings.sorting[thPane.dataset.name] = order;

				this.paneContent.querySelector('.table-pane').innerHTML = this.generatePaneContent({ table: newContent.querySelector('.table') }).innerHTML;

				this.paneContent.querySelector('thead').querySelectorAll('th').forEach((_thPane, position) => {
					tableHeadHandler(_thPane, table.querySelector('thead').querySelectorAll('th')[position]);
				});
			});

			thPane.querySelector('input').onChanged(value => {
				let name = 'crater-table-data-' + value.toLowerCase();
				let ths = this.paneContent.querySelector('thead').querySelectorAll('th');

				for (let sibling of ths) {
					if (sibling != thPane && sibling.dataset.name == name) {
						alert('Column already exists, Try another name');
						return;
					}
				}

				let tds = this.paneContent.querySelectorAll('td');

				for (let i in tds) {
					let td = tds[i];
					if (td.nodeName == 'TD' && td.dataset.name == thPane.dataset.name) {
						td.dataset.name = name;
						table.querySelectorAll('td')[i].dataset.name = name;
					}
				}

				thDom.textContent = value;
				thDom.dataset.name = name;
				thPane.dataset.name = name;
			});
		};

		this.paneContent.querySelector('thead').querySelectorAll('th').forEach((thPane, position) => {
			tableHeadHandler(thPane, table.querySelector('thead').querySelectorAll('th')[position]);
		});

		let tableBodyDataHandler = (td) => {
			td.addEventListener('mouseover', event => {
				for (let th of this.paneContent.querySelector('thead').querySelectorAll('th')) {
					if (th.dataset.name == td.dataset.name) {
						th.querySelector('.crater-content-options').css({ visibility: 'visible' });
					}
				}
			});

			td.addEventListener('mouseout', event => {
				for (let th of this.paneContent.querySelector('thead').querySelectorAll('th')) {
					if (th.dataset.name == td.dataset.name) {
						th.querySelector('.crater-content-options').css({ visibility: 'hidden' });
					}
				}
			});
		};

		this.paneContent.querySelector('tbody').querySelectorAll('td').forEach(td => {
			tableBodyDataHandler(td);
		});

		let getName = () => {
			let otherThs = this.paneContent.querySelectorAll('th');
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

		this.paneContent.querySelector('thead').addEventListener('click', event => {
			let target = event.target;
			if (target.classList.contains('delete-crater-table-content-column')) {
				if (!confirm("Do you really want to delete this column")) {
					return;
				}

				let th = target.getParents('TH');

				let name = th.dataset.name;

				//remove the tds
				this.paneContent.querySelectorAll('td').forEach((td) => {
					if (td.nodeName == 'TD' && td.dataset.name == name) {
						td.remove();
					}
				});

				table.querySelectorAll('td').forEach((td) => {
					if (td.nodeName == 'TD' && td.dataset.name == name) {
						td.remove();
					}
				});

				//remove the TH
				table.querySelector('thead').querySelectorAll('th').forEach(aTH => {
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
				this.paneContent.querySelectorAll('td').forEach((td) => {
					if (td.nodeName == 'TD' && td.dataset.name == th.dataset.name) {
						let aTDClone = td.cloneNode(true);
						aTDClone.dataset.name = copyName;
						td.before(aTDClone);
						tableBodyDataHandler(aTDClone);
					}
				});

				table.querySelectorAll('td').forEach((td) => {
					if (td.nodeName == 'TD' && td.dataset.name == th.dataset.name) {
						let aTDClone = td.cloneNode(true);
						aTDClone.dataset.name = copyName;
						td.before(aTDClone);
					}
				});

				//remove the TH
				let newPaneTH: any;
				table.querySelector('thead').querySelectorAll('th').forEach(aTH => {
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
				aTHPaneClone.querySelector('input').setAttribute('value', `${'SAMPLE'}${copyName.slice(dataName.length)}`);
				th.before(aTHPaneClone);
				tableHeadHandler(aTHPaneClone, newPaneTH);
			}
			else if (target.classList.contains('add-after-crater-table-content-column')) {
				let th = target.getParents('TH');
				let copyName = getName();
				//remove the tds
				this.paneContent.querySelectorAll('td').forEach((td) => {
					if (td.nodeName == 'TD' && td.dataset.name == th.dataset.name) {
						let aTDClone = td.cloneNode(true);
						aTDClone.dataset.name = copyName;
						td.after(aTDClone);
						tableBodyDataHandler(aTDClone);
					}
				});

				table.querySelectorAll('td').forEach((td) => {
					if (td.nodeName == 'TD' && td.dataset.name == th.dataset.name) {
						let aTDClone = td.cloneNode(true);
						aTDClone.dataset.name = copyName;
						td.after(aTDClone);
					}
				});

				//remove the TH
				let newPaneTH: any;
				table.querySelector('thead').querySelectorAll('th').forEach(aTH => {
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
				aTHPaneClone.querySelector('input').setAttribute('value', `${'SAMPLE'}${copyName.slice(dataName.length)}`);
				th.after(aTHPaneClone);
				tableHeadHandler(aTHPaneClone, newPaneTH);
			}
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());

			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart            
		});

		let tableSettings = this.paneContent.querySelector('.table-settings');
		let tableHeaderSettings = this.paneContent.querySelector('.table-header-settings');
		let tableBodyDataSettings = this.paneContent.querySelector('.table-data-settings');

		tableHeaderSettings.querySelector('#fontsize-cell').onChanged(value => {
			table.querySelectorAll('th').forEach(th => {
				th.css({ fontSize: value });
			});
		});

		tableHeaderSettings.querySelector('#show-cell').onChanged(value => {
			if (value == 'No') {
				table.querySelector('thead').hide();
			} else {
				table.querySelector('thead').show();
			}
		});

		let headerColorCell = tableHeaderSettings.querySelector('#color-cell').parentNode;
		this.pickColor({ parent: headerColorCell, cell: headerColorCell.querySelector('#color-cell') }, (color) => {
			table.querySelectorAll('th').forEach(th => {
				th.css({ color });
			});
			headerColorCell.querySelector('#color-cell').value = color;
			headerColorCell.querySelector('#color-cell').setAttribute('value', color);
		});

		let headerBackgroundColorCell = tableHeaderSettings.querySelector('#backgroundcolor-cell').parentNode;
		this.pickColor({ parent: headerBackgroundColorCell, cell: headerBackgroundColorCell.querySelector('#backgroundcolor-cell') }, (backgroundColor) => {
			table.querySelector('thead').querySelector('tr').css({ backgroundColor });
			headerBackgroundColorCell.querySelector('#backgroundcolor-cell').value = backgroundColor;
			headerBackgroundColorCell.querySelector('#backgroundcolor-cell').setAttribute('value', backgroundColor);
		});

		tableHeaderSettings.querySelector('#height-cell').onChanged(value => {
			table.querySelector('thead').querySelector('tr').css({ height: value });
		});

		tableBodyDataSettings.querySelector('#fontsize-cell').onChanged(value => {
			table.querySelectorAll('td').forEach(th => {
				th.css({ fontSize: value });
			});
		});

		let dataColorCell = tableBodyDataSettings.querySelector('#color-cell').parentNode;
		this.pickColor({ parent: dataColorCell, cell: dataColorCell.querySelector('#color-cell') }, (color) => {
			table.querySelectorAll('td').forEach(td => {
				td.css({ color });
			});
			dataColorCell.querySelector('#color-cell').value = color;
			dataColorCell.querySelector('#color-cell').setAttribute('value', color);
		});

		let dataBackgroundColorCell = tableBodyDataSettings.querySelector('#backgroundcolor-cell').parentNode;
		this.pickColor({ parent: dataBackgroundColorCell, cell: dataBackgroundColorCell.querySelector('#backgroundcolor-cell') }, (backgroundColor) => {
			table.querySelector('tbody').querySelectorAll('tr').forEach(tr => {
				tr.css({ backgroundColor });
			});
			dataBackgroundColorCell.querySelector('#backgroundcolor-cell').value = backgroundColor;
			dataBackgroundColorCell.querySelector('#backgroundcolor-cell').setAttribute('value', backgroundColor);
		});

		tableBodyDataSettings.querySelector('#backgroundcolor-cell').onChanged(value => {
			table.querySelector('tbody').querySelectorAll('tr').forEach(tr => {
				tr.css({ backgroundColor: value });
			});
		});

		tableBodyDataSettings.querySelector('#height-cell').onChanged(value => {
			table.querySelector('tbody').querySelectorAll('tr').forEach(tr => {
				tr.css({ height: value });
			});
		});

		tableSettings.querySelector('#bordersize-cell').onChanged(value => {
			table.querySelectorAll('tr').forEach(tr => {
				tr.css({ borderWidth: value });
			});
		});

		let borderColorCell = tableSettings.querySelector('#bordercolor-cell').parentNode;
		this.pickColor({ parent: borderColorCell, cell: borderColorCell.querySelector('#bordercolor-cell') }, (borderColor) => {
			table.querySelectorAll('tr').forEach(tr => {
				tr.css({ borderColor });
			});
			borderColorCell.querySelector('#bordercolor-cell').value = borderColor;
			borderColorCell.querySelector('#bordercolor-cell').setAttribute('value', borderColor);
		});

		tableSettings.querySelector('#borderstyle-cell').onChanged(value => {
			table.querySelectorAll('tr').forEach(tr => {
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

		let paneConnection = this.sharePoint.app.querySelector('.crater-property-connection');

		let metaWindow = this.elementModifier.createForm({
			title: 'Set Table Sample', attributes: { id: 'meta-form', class: 'form' },
			contents: {
				Names: { element: 'input', attributes: { id: 'meta-data-names', name: 'Names', value: headers }, options: params.options, note: 'Names of data should be comma seperated[data1, data2]' },
			},
			buttons: {
				submit: { element: 'button', attributes: { id: 'set-meta', class: 'btn' }, text: 'Set' },
			}
		});

		metaWindow.querySelector('#set-meta').addEventListener('click', event => {
			event.preventDefault();

			let names = metaWindow.querySelector('#meta-data-names').value.split(',');
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

			updateWindow.querySelector('#update-element').addEventListener('click', updateEvent => {
				event.preventDefault();
				let formData = updateWindow.querySelectorAll('.form-data');

				for (let i = 0; i < formData.length; i++) {
					data[formData[i].name.toLowerCase()] = formData[i].value;
				}

				source = func.extractFromJsonArray(data, params.source);

				let newContent = this.render({ source });
				draftDom.querySelector('.crater-table').innerHTML = newContent.querySelector('.crater-table').innerHTML;

				this.sharePoint.properties.pane.content[this.key].draft.html = draftDom.outerHTML;

				this.paneContent.querySelector('.table-pane').innerHTML = this.generatePaneContent({ table: newContent.querySelector('table') }).innerHTML;

				this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			});

			let parent = metaWindow.parentNode;
			parent.innerHTML = '';
			parent.append(metaWindow, updateWindow);
		});

		if (!func.isnull(paneConnection)) {
			paneConnection.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {

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
								element: 'input', name: 'title', value: this.element.querySelector('.crater-panel-title').textContent
							}),
							this.elementModifier.cell({
								element: 'input', name: 'backgroundcolor', value: this.element.querySelector('.crater-panel-title').css()['background-color'], list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'color', value: this.element.querySelector('.crater-panel-title').css().color, list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontsize', value: this.element.querySelector('.crater-panel-title').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'height', value: this.element.querySelector('.crater-panel-title').css()['height']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'width', value: this.element.querySelector('.crater-panel-title').css()['width']
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

			let contents = this.element.querySelector('.crater-panel-content').querySelectorAll('.keyed-element');

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
		this.paneContent = this.sharePoint.app.querySelector('.crater-property-content').monitor();
		let panelContents = this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-panel-content');
		let panelContentDom = panelContents.childNodes;
		let panelContentPane = this.paneContent.querySelector('.panel-contents-pane');

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

			panelContentRowPane.querySelector('#name').innerHTML = panelContentRowDom.dataset.type.toUpperCase();

			panelContentRowPane.addEventListener('mouseover', event => {
				panelContentRowPane.querySelector('.crater-content-options').css({ visibility: 'visible' });
			});

			panelContentRowPane.addEventListener('mouseout', event => {
				panelContentRowPane.querySelector('.crater-content-options').css({ visibility: 'hidden' });
			});

			panelContentRowPane.querySelector('.delete-crater-panel-content-row').addEventListener('click', event => {
				panelContentRowDom.remove();
				panelContentRowPane.remove();
			});

			panelContentRowPane.querySelector('.add-before-crater-panel-content-row').addEventListener('click', event => {
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

			panelContentRowPane.querySelector('.add-after-crater-panel-content-row').addEventListener('click', event => {
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

		let titlePane = this.paneContent.querySelector('.title-pane');
		let titleLinkPane = this.paneContent.querySelector('.title-link-pane');
		let settingsPane = this.paneContent.querySelector('.settings-pane');

		let title = this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-panel-title');
		titlePane.querySelector('#height-cell').onChanged(height => {
			title.css({ height });
		});

		titlePane.querySelector('#width-cell').onChanged(width => {
			title.css({ width });
		});

		titlePane.querySelector('#position-cell').onChanged(position => {
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

		titlePane.querySelector('#layout-cell').onChanged(layout => {
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

		titlePane.querySelector('#title-cell').onChanged(value => {
			title.querySelector('.crater-panel-title-text').innerText = value;
		});

		titleLinkPane.querySelector('#text-cell').onChanged(value => {
			title.querySelector('.crater-panel-title-link').innerText = value;
		});

		titleLinkPane.querySelector('#color-cell').onChanged(color => {
			title.querySelector('.crater-panel-title-link').css({ color });
		});

		titleLinkPane.querySelector('#background-color-cell').onChanged(backgroundColor => {
			title.querySelector('.crater-panel-title-link').css({ backgroundColor });
		});

		titleLinkPane.querySelector('#border-cell').onChanged(border => {
			title.querySelector('.crater-panel-title-link').css({ border });
		});

		titleLinkPane.querySelector('#url-cell').onChanged(value => {
			title.querySelector('.crater-panel-title-link').href = value;
		});

		titleLinkPane.querySelector('#Show-cell').onChanged(value => {
			if (value == 'No') {
				title.querySelector('.crater-panel-title-link').css({ display: 'none' });
			}
			else {
				title.querySelector('.crater-panel-title-link').cssRemove(['display']);
			}
		});

		let backgroundColorCell = titlePane.querySelector('#backgroundcolor-cell').parentNode;
		this.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.querySelector('#backgroundcolor-cell') }, (backgroundColor) => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-panel-title').css({ backgroundColor });
			backgroundColorCell.querySelector('#backgroundcolor-cell').value = backgroundColor;
			backgroundColorCell.querySelector('#backgroundcolor-cell').setAttribute('value', backgroundColor);
			this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-panel-content').css({
				borderColor: backgroundColor
			});
		});

		let colorCell = titlePane.querySelector('#color-cell').parentNode;
		this.pickColor({ parent: colorCell, cell: colorCell.querySelector('#color-cell') }, (color) => {
			title.querySelector('.crater-panel-title-text').css({ color });
			colorCell.querySelector('#color-cell').value = color;
			colorCell.querySelector('#color-cell').setAttribute('value', color);
		});

		titlePane.querySelector('#fontsize-cell').onChanged(value => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-panel-title').css({ fontSize: value });
		});

		settingsPane.querySelector('#Box-Content-cell').onChanged(value => {
			if (value == 'Yes') {
				this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-panel-content').css({
					borderColor: title.css().backgroundColor
				});
			} else {
				this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-panel-content').css({
					borderColor: 'transparent'
				});
			}
		});

		this.paneContent.querySelector('.new-component').addEventListener('click', event => {

			this.sharePoint.app.querySelectorAll('.crater-display-panel').forEach(panel => {
				panel.remove();
			});

			this.paneContent.append(this.sharePoint.displayPanel(webpart => {
				let newPanelContent = this.sharePoint.appendWebpart(panelContents, webpart.dataset.webpart);
				let newPanelContentRow = panelContentPanePrototype.cloneNode(true);
				panelContentPane.append(newPanelContentRow);

				panelContentRowHandler(newPanelContentRow, newPanelContent);
			}));
		});

		this.paneContent.querySelectorAll('.crater-panel-content-row-pane').forEach((panelContent, position) => {
			panelContentRowHandler(panelContent, panelContentDom[position]);
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;//update webpart
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;

			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());

			for (let keyedElement of this.element.querySelectorAll('.keyed-element')) {
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

		this.element.querySelector('.crater-countdown-days').querySelector('.crater-countdown-counting').innerText = date.days;

		this.element.querySelector('.crater-countdown-hours').querySelector('.crater-countdown-counting').innerText = date.hours;

		this.element.querySelector('.crater-countdown-minutes').querySelector('.crater-countdown-counting').innerText = date.minutes;

		this.element.querySelector('.crater-countdown-seconds').querySelector('.crater-countdown-counting').innerText = date.seconds;

		this.sharePoint.properties.pane.content[this.key].settings.interval = setInterval(() => {
			date = this.getDate(secondsTogoCurrently);

			if (date.past) {
				this.element.classList.toggle('crater-countdown-past');
			} else {
				this.element.classList.remove('crater-countdown-past');
			}

			this.element.querySelector('.crater-countdown-days').querySelector('.crater-countdown-counting').innerText = date.days;

			this.element.querySelector('.crater-countdown-hours').querySelector('.crater-countdown-counting').innerText = date.hours;

			this.element.querySelector('.crater-countdown-minutes').querySelector('.crater-countdown-counting').innerText = date.minutes;

			this.element.querySelector('.crater-countdown-seconds').querySelector('.crater-countdown-counting').innerText = date.seconds;
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
		this.paneContent = this.sharePoint.app.querySelector('.crater-property-content').monitor();

		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;

		let countingPane = this.paneContent.querySelector('.counting-pane');

		let countingColorCell = countingPane.querySelector('#Color-cell').parentNode;
		this.pickColor({ parent: countingColorCell, cell: countingColorCell.querySelector('#Color-cell') }, (color) => {
			draftDom.querySelectorAll('.crater-countdown-counting').forEach(element => {
				element.css({ color });
			});
			countingColorCell.querySelector('#Color-cell').value = color;
			countingColorCell.querySelector('#Color-cell').setAttribute('value', color);
		});

		let countingBackgroundColorCell = countingPane.querySelector('#BackgroundColor-cell').parentNode;
		this.pickColor({ parent: countingBackgroundColorCell, cell: countingBackgroundColorCell.querySelector('#BackgroundColor-cell') }, (backgroundColor) => {
			draftDom.querySelectorAll('.crater-countdown-counting').forEach(element => {
				element.css({ backgroundColor });
			});
			countingBackgroundColorCell.querySelector('#BackgroundColor-cell').value = backgroundColor;
			countingBackgroundColorCell.querySelector('#BackgroundColor-cell').setAttribute('value', backgroundColor);
		});

		countingPane.querySelector('#FontSize-cell').onChanged(fontSize => {
			draftDom.querySelectorAll('.crater-countdown-counting').forEach(element => {
				element.css({ fontSize });
			});
		});

		countingPane.querySelector('#FontStyle-cell').onChanged(fontFamily => {
			draftDom.querySelectorAll('.crater-countdown-counting').forEach(element => {
				element.css({ fontFamily });
			});
		});

		let labelPane = this.paneContent.querySelector('.label-pane');

		let labelColorCell = labelPane.querySelector('#Color-cell').parentNode;
		this.pickColor({ parent: labelColorCell, cell: labelColorCell.querySelector('#Color-cell') }, (color) => {
			draftDom.querySelectorAll('.crater-countdown-label').forEach(element => {
				element.css({ color });
			});
			labelColorCell.querySelector('#Color-cell').value = color;
			labelColorCell.querySelector('#Color-cell').setAttribute('value', color);
		});

		let labelBackgroundColorCell = labelPane.querySelector('#BackgroundColor-cell').parentNode;
		this.pickColor({ parent: labelBackgroundColorCell, cell: labelBackgroundColorCell.querySelector('#BackgroundColor-cell') }, (backgroundColor) => {
			draftDom.querySelectorAll('.crater-countdown-label').forEach(element => {
				element.css({ backgroundColor });
			});
			labelBackgroundColorCell.querySelector('#BackgroundColor-cell').value = backgroundColor;
			labelBackgroundColorCell.querySelector('#BackgroundColor-cell').setAttribute('value', backgroundColor);
		});

		labelPane.querySelector('#FontSize-cell').onChanged(size => {
			draftDom.querySelectorAll('.crater-countdown-label').forEach(element => {
				element.css({ fontSize: size });
			});
		});

		labelPane.querySelector('#FontStyle-cell').onChanged(style => {
			draftDom.querySelectorAll('.crater-countdown-label').forEach(element => {
				element.css({ fontFamily: style });
			});
		});

		let settingsPane = this.paneContent.querySelector('.settings-pane');

		let settingsDate = settingsPane.querySelector('#Date-cell');
		let settingsTime = settingsPane.querySelector('#Time-cell');
		let settingsBorder = settingsPane.querySelector('#Border-cell');
		let settingsBorderRadius = settingsPane.querySelector('#BorderRadius-cell');

		settingsDate.onChanged(date => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.dataset.date = func.dateValue(date);
		});

		settingsTime.onChanged(time => {
			if (func.isTimeValid(time)) {
				this.sharePoint.properties.pane.content[this.key].draft.dom.dataset.time = func.isTimeValid(time);
			}
		});

		settingsBorder.onChanged(border => {
			draftDom.querySelectorAll('.crater-countdown-block').forEach(element => {
				element.css({ border });
			});
		});

		settingsBorderRadius.onChanged(borderRadius => {
			draftDom.querySelectorAll('.crater-countdown-block').forEach(element => {
				element.css({ borderRadius });
			});
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
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

		let dateListContent = dateList.querySelector(`.crater-datelist-content`);


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
			let dateList = this.sharePoint.properties.pane.content[key].draft.dom.querySelector('.crater-datelist-content');
			let dateListRows = dateList.querySelectorAll('.crater-datelist-content-item');
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
								element: 'img', name: 'icon', attributes: {}, dataAttributes: { class: 'crater-icon', src: this.element.querySelector('.crater-datelist-title-imgIcon').src }
							}),
							this.elementModifier.cell({
								element: 'input', name: 'title', value: this.element.querySelector('.crater-datelist-title').textContent
							}),
							this.elementModifier.cell({
								element: 'input', name: 'backgroundcolor', value: this.element.querySelector('.crater-datelist-title').css()['background-color'], list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'color', value: this.element.querySelector('.crater-datelist-title').css().color, list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontsize', value: this.element.querySelector('.crater-datelist-title').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'height', value: this.element.querySelector('.crater-datelist-title').css()['height'] || ''
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
								element: 'input', name: 'daySize', value: this.element.querySelector('.crater-datelist-content-item-date-day').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'monthSize', value: this.element.querySelector('.crater-datelist-content-item-date-month').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontFamily', value: this.element.querySelector('.crater-datelist-content-item-date').css()['font-family']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'dayColor', value: this.element.querySelector('.crater-datelist-content-item-date-day').css()['color'], list: func.colors
							}),
							this.elementModifier.cell({
								element: 'input', name: 'monthColor', value: this.element.querySelector('.crater-datelist-content-item-date-month').css()['color'], list: func.colors
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
								element: 'input', name: 'fontSize', value: this.element.querySelector('.crater-datelist-content-item-text-main').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontFamily', value: this.element.querySelector('.crater-datelist-content-item-text-main').css()['font-family']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'titleColor', value: this.element.querySelector('.crater-datelist-content-item-text-main').css()['color'], list: func.colors
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
								element: 'input', name: 'fontSize', value: this.element.querySelector('.crater-datelist-content-item-text-subtitle').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontFamily', value: this.element.querySelector('.crater-datelist-content-item-text-subtitle').css()['font-family']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'subtitleColor', value: this.element.querySelector('.crater-datelist-content-item-text-subtitle').css().color, list: func.colors
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
								element: 'input', name: 'fontSize', value: this.element.querySelector('.crater-datelist-content-item-text-body').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontFamily', value: this.element.querySelector('.crater-datelist-content-item-text-body').css()['font-family']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'bodyColor', value: this.element.querySelector('.crater-datelist-content-item-text-body').css()['color'], list: func.colors
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
						element: 'input', name: 'Day', attribute: { class: 'crater-date' }, value: params.list[i].querySelector('#Day').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Month', attribute: { class: 'crater-date' }, value: params.list[i].querySelector('#Month').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'title', value: params.list[i].querySelector('#mainText').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'subtitle', value: params.list[i].querySelector('#subtitle').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'body', value: params.list[i].querySelector('#body').textContent
					}),
				]
			});
		}

		return dateListPane;

	}

	public listenPaneContent(params) {
		this.element = params.element;
		this.key = this.element.dataset['key'];
		this.paneContent = this.sharePoint.app.querySelector('.crater-property-content').monitor();
		//get the content and all the events
		let dateList = this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-datelist-content');
		let dateListRow = dateList.querySelectorAll('.crater-datelist-content-item');

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
				dateRowPane.querySelector('.crater-content-options').css({ visibility: 'visible' });
			});

			dateRowPane.addEventListener('mouseout', event => {
				dateRowPane.querySelector('.crater-content-options').css({ visibility: 'hidden' });
			});

			// get the values of the newly created row on the property - pane
			dateRowPane.querySelector('#title-cell').onChanged(value => {
				dateRowDom.querySelector('.crater-datelist-content-item-text-main').innerHTML = value;
			});

			dateRowPane.querySelector('#subtitle-cell').onChanged(value => {
				dateRowDom.querySelector('.crater-datelist-content-item-text-subtitle').innerHTML = value;
			});

			dateRowPane.querySelector('#body-cell').onChanged(value => {
				dateRowDom.querySelector('.crater-datelist-content-item-text-body').innerHTML = value;
			});

			dateRowPane.querySelector('#Day-cell').onChanged(value => {
				dateRowDom.querySelector('.crater-datelist-content-item-date-day').innerHTML = value;
			});

			dateRowPane.querySelector('#Month-cell').onChanged(value => {
				dateRowDom.querySelector('.crater-datelist-content-item-date-month').innerHTML = value;
			});

			dateRowPane.querySelector('.delete-crater-datelist-content-item').addEventListener('click', event => {
				dateRowDom.remove();
				dateRowPane.remove();
			});

			dateRowPane.querySelector('.add-before-crater-datelist-content-item').addEventListener('click', event => {
				let newdateRowDom = dateListRowDomPrototype.cloneNode(true);
				let newdateRowPane = dateListRowPanePrototype.cloneNode(true);

				dateRowDom.before(newdateRowDom);
				dateRowPane.before(newdateRowPane);
				dateRowHandler(newdateRowPane, newdateRowDom);
			});

			dateRowPane.querySelector('.add-after-crater-datelist-content-item').addEventListener('click', event => {
				let newdateRowDom = dateListRowDomPrototype.cloneNode(true);
				let newdateRowPane = dateListRowPanePrototype.cloneNode(true);

				dateRowDom.after(newdateRowDom);
				dateRowPane.after(newdateRowPane);

				dateRowHandler(newdateRowPane, newdateRowDom);
			});
		};

		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;

		let titlePane = this.paneContent.querySelector('.title-pane');
		let dateListDateRowPane = this.paneContent.querySelector('.datelist-date-row-pane');
		let dateListTitleRowPane = this.paneContent.querySelector('.datelist-title-row-pane');
		let dateListSubtitleRowPane = this.paneContent.querySelector('.datelist-subtitle-row-pane');
		let dateListBodyRowPane = this.paneContent.querySelector('.datelist-body-row-pane');

		let dateListTitleParent = dateListTitleRowPane.querySelector('#titleColor-cell').parentNode;
		this.pickColor({ parent: dateListTitleParent, cell: dateListTitleParent.querySelector('#titleColor-cell') }, (color) => {
			draftDom.querySelectorAll('.crater-datelist-content-item-text-main').forEach(element => {
				element.css({ color });//get the color of the event font in the draftDom
			});
			dateListTitleParent.querySelector('#titleColor-cell').value = color;
			dateListTitleParent.querySelector('#titleColor-cell').setAttribute('value', color); //set the value of the eventColor cell in the pane to the color
		});

		let dateListSubtitleParent = dateListSubtitleRowPane.querySelector('#subtitleColor-cell').parentNode;
		this.pickColor({ parent: dateListSubtitleParent, cell: dateListSubtitleParent.querySelector('#subtitleColor-cell') }, (color) => {
			draftDom.querySelectorAll('.crater-datelist-content-item-text-subtitle').forEach(element => {
				element.css({ color });//get the color of the event font in the draftDom
			});
			dateListSubtitleParent.querySelector('#subtitleColor-cell').value = color;
			dateListSubtitleParent.querySelector('#subtitleColor-cell').setAttribute('value', color); //set the value of the eventColor cell in the pane to the color
		});

		let dateListBodyParent = dateListBodyRowPane.querySelector('#bodyColor-cell').parentNode;
		this.pickColor({ parent: dateListBodyParent, cell: dateListBodyParent.querySelector('#bodyColor-cell') }, (color) => {
			draftDom.querySelectorAll('.crater-datelist-content-item-text-body').forEach(element => {
				element.css({ color });//get the color of the event font in the draftDom
			});
			dateListBodyParent.querySelector('#bodyColor-cell').value = color;
			dateListBodyParent.querySelector('#bodyColor-cell').setAttribute('value', color); //set the value of the eventColor cell in the pane to the color
		});

		let dateListDayParent = dateListDateRowPane.querySelector('#dayColor-cell').parentNode;
		this.pickColor({ parent: dateListDayParent, cell: dateListDayParent.querySelector('#dayColor-cell') }, (color) => {
			draftDom.querySelectorAll('.crater-datelist-content-item-date-day').forEach(element => {
				element.css({ color });//get the color of the event font in the draftDom
			});
			dateListDayParent.querySelector('#dayColor-cell').value = color;
			dateListDayParent.querySelector('#dayColor-cell').setAttribute('value', color); //set the value of the eventColor cell in the pane to the color
		});

		let dateListMonthParent = dateListDateRowPane.querySelector('#monthColor-cell').parentNode;
		this.pickColor({ parent: dateListMonthParent, cell: dateListMonthParent.querySelector('#monthColor-cell') }, (color) => {
			draftDom.querySelectorAll('.crater-datelist-content-item-date-month').forEach(element => {
				element.css({ color });//get the color of the event font in the draftDom
			});
			dateListMonthParent.querySelector('#monthColor-cell').value = color;
			dateListMonthParent.querySelector('#monthColor-cell').setAttribute('value', color); //set the value of the eventColor cell in the pane to the color
		});

		let iconParent = titlePane.querySelector('#icon-cell').parentNode;
		this.uploadImage({ parent: iconParent }, (image) => {
			iconParent.querySelector('#icon-cell').src = image.src;
			draftDom.querySelector('.crater-datelist-title-imgIcon').src = image.src;
		});

		titlePane.querySelector('#title-cell').onChanged(value => {
			draftDom.querySelector('.crater-datelist-title-captionTitle').innerHTML = value;
		});

		let backgroundColorCell = titlePane.querySelector('#backgroundcolor-cell').parentNode;
		this.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.querySelector('#backgroundcolor-cell') }, (backgroundColor) => {
			draftDom.querySelector('.crater-datelist-title').css({ backgroundColor });
			backgroundColorCell.querySelector('#backgroundcolor-cell').value = backgroundColor;
			backgroundColorCell.querySelector('#backgroundcolor-cell').setAttribute('value', backgroundColor);
		});

		let colorCell = titlePane.querySelector('#color-cell').parentNode;
		this.pickColor({ parent: colorCell, cell: colorCell.querySelector('#color-cell') }, (color) => {
			draftDom.querySelector('.crater-datelist-title').css({ color });
			colorCell.querySelector('#color-cell').value = color;
			colorCell.querySelector('#color-cell').setAttribute('value', color);
		});


		titlePane.querySelector('#fontsize-cell').onChanged(value => {
			draftDom.querySelector('.crater-datelist-title').css({ fontSize: value });
		});

		titlePane.querySelector('#height-cell').onChanged(value => {
			draftDom.querySelector('.crater-datelist-title').css({ height: value });
		});



		dateListTitleRowPane.querySelector('#fontSize-cell').onChanged(value => {
			draftDom.querySelectorAll('.crater-datelist-content-item-text-main').forEach(element => {
				element.css({ fontSize: value });
			});
		});
		dateListTitleRowPane.querySelector('#fontFamily-cell').onChanged(value => {
			draftDom.querySelectorAll('.crater-datelist-content-item-text-main').forEach(element => {
				element.css({ fontFamily: value });
			});
		});

		dateListSubtitleRowPane.querySelector('#fontSize-cell').onChanged(value => {
			draftDom.querySelectorAll('.crater-datelist-content-item-text-subtitle').forEach(element => {
				element.css({ fontSize: value });
			});
		});
		dateListSubtitleRowPane.querySelector('#fontFamily-cell').onChanged(value => {
			draftDom.querySelectorAll('.crater-datelist-content-item-text-subtitle').forEach(element => {
				element.css({ fontFamily: value });
			});
		});

		dateListBodyRowPane.querySelector('#fontSize-cell').onChanged(value => {
			draftDom.querySelectorAll('.crater-datelist-content-item-text-body').forEach(element => {
				element.css({ fontSize: value });
			});
		});
		dateListBodyRowPane.querySelector('#fontFamily-cell').onChanged(value => {
			draftDom.querySelectorAll('.crater-datelist-content-item-text-body').forEach(element => {
				element.css({ fontFamily: value });
			});
		});

		dateListDateRowPane.querySelector('#daySize-cell').onChanged(value => {
			draftDom.querySelectorAll('.crater-datelist-content-item-date-day').forEach(element => {
				element.css({ fontSize: value });
			});
		});
		dateListDateRowPane.querySelector('#monthSize-cell').onChanged(value => {
			draftDom.querySelectorAll('.crater-datelist-content-item-date-month').forEach(element => {
				element.css({ fontSize: value });
			});
		});

		dateListDateRowPane.querySelector('#fontFamily-cell').onChanged(value => {
			draftDom.querySelectorAll('.crater-datelist-content-item-date').forEach(element => {
				element.css({ fontFamily: value });
			});
		});

		//appends the dom and pane prototypes to the dom and pane when you click add new
		this.paneContent.querySelector('.new-component').addEventListener('click', event => {
			let newDateRowDom = dateListRowDomPrototype.cloneNode(true);
			let newDateRowPane = dateListRowPanePrototype.cloneNode(true);

			dateList.append(newDateRowDom);//c
			this.paneContent.querySelector('.datelist-pane').append(newDateRowPane);
			dateRowHandler(newDateRowPane, newDateRowDom);
		});

		let paneItems = this.paneContent.querySelectorAll('.crater-datelist-item-pane');
		paneItems.forEach((dateRow, position) => {
			dateRowHandler(dateRow, dateListRow[position]);
		});

		let showHeader = titlePane.querySelector('#toggleTitle-cell');
		showHeader.addEventListener('change', e => {

			switch (showHeader.value.toLowerCase()) {
				case "hide":
					draftDom.querySelectorAll('.crater-datelist-title').forEach(element => {
						element.style.display = "none";
					});
					break;
				case "show":
					draftDom.querySelectorAll('.crater-datelist-title').forEach(element => {
						element.style.display = "grid";
					});
					break;
				default:
					draftDom.querySelectorAll('.crater-datelist-title').forEach(element => {
						element.style.display = "none";
					});
			}
		});

		let showTitle = dateListTitleRowPane.querySelector('#toggleTitle-cell');
		showTitle.addEventListener('change', e => {

			switch (showTitle.value.toLowerCase()) {
				case "hide":
					draftDom.querySelectorAll('.crater-datelist-content-item-text-main').forEach(element => {
						element.style.display = "none";
					});
					break;
				case "show":
					draftDom.querySelectorAll('.crater-datelist-content-item-text-main').forEach(element => {
						element.style.display = "block";
					});
					break;
				default:
					draftDom.querySelectorAll('.crater-datelist-content-item-text-main').forEach(element => {
						element.style.display = "none";
					});
			}
		});

		let showSubtitle = dateListSubtitleRowPane.querySelector('#toggleSubtitle-cell');
		showSubtitle.addEventListener('change', e => {

			switch (showSubtitle.value.toLowerCase()) {
				case "hide":
					draftDom.querySelectorAll('.crater-datelist-content-item-text-subtitle').forEach(element => {
						element.style.display = "none";
					});
					break;
				case "show":
					draftDom.querySelectorAll('.crater-datelist-content-item-text-subtitle').forEach(element => {
						element.style.display = "block";
					});
					break;
				default:
					draftDom.querySelectorAll('.crater-datelist-content-item-text-subtitle').forEach(element => {
						element.style.display = "none";
					});
			}
		});

		let showBody = dateListBodyRowPane.querySelector('#toggleBody-cell');
		showBody.addEventListener('change', e => {

			switch (showBody.value.toLowerCase()) {
				case "hide":
					draftDom.querySelectorAll('.crater-datelist-content-item-text-body').forEach(element => {
						element.style.display = "none";
					});
					break;
				case "show":
					draftDom.querySelectorAll('.crater-datelist-content-item-text-body').forEach(element => {
						element.style.display = "block";
					});
					break;
				default:
					draftDom.querySelectorAll('.crater-datelist-content-item-text-body').forEach(element => {
						element.style.display = "none";
					});
			}
		});

		let showDate = dateListDateRowPane.querySelector('#toggleDate-cell');
		showDate.addEventListener('change', e => {

			switch (showDate.value.toLowerCase()) {
				case "hide":
					draftDom.querySelectorAll('.crater-datelist-content-item-date').forEach(element => {
						element.style.visibility = "hidden";
					});
					break;
				case "show":
					draftDom.querySelectorAll('.crater-datelist-content-item-date').forEach(element => {
						element.style.visibility = "visible";
					});
					break;
				default:
					draftDom.querySelectorAll('.crater-datelist-content-item-date').forEach(element => {
						element.style.visibility = "hidden";
					});
			}
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
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

		let paneConnection = this.sharePoint.app.querySelector('.crater-property-connection');

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
		updateWindow.querySelector('#update-element').addEventListener('click', event => {
			event.preventDefault();
			data.day = updateWindow.querySelector('#meta-data-day').value;
			data.month = updateWindow.querySelector('#meta-data-month').value;
			data.title = updateWindow.querySelector('#meta-data-title').value;
			data.subtitle = updateWindow.querySelector('#meta-data-subtitle').value;
			data.body = updateWindow.querySelector('#meta-data-body').value;
			source = func.extractFromJsonArray(data, params.source);

			let newContent = this.render({ source });
			draftDom.querySelector('.crater-datelist-content').innerHTML = newContent.querySelector('.crater-datelist-content').innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = draftDom.outerHTML;
			this.paneContent.querySelector('.datelist-pane').innerHTML = this.generatePaneContent({ list: newContent.querySelectorAll('.crater-datelist-content-item') }).innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
		});


		if (!func.isnull(paneConnection)) {
			paneConnection.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
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
		let map = new google.maps.Map(this.element.querySelector('#crater-map-div'), {
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
		this.element.querySelector('#crater-map-div').innerHTML = '';
		this.element.querySelector('script').remove();
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
								element: 'input', name: 'width', value: this.element.querySelector('#crater-map-div').css()['width'], list: func.pixelSizes
							}),
							this.elementModifier.cell({
								element: 'input', name: 'height', value: this.element.querySelector('#crater-map-div').css()['height'], list: func.pixelSizes
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

		this.paneContent = this.sharePoint.app.querySelector('.crater-property-content').monitor();

		let mapPane = this.paneContent.querySelector('.map-style-pane');

		mapPane.querySelector('#latitude-cell').onChanged(value => {
			this.sharePoint.properties.pane.content[this.key].settings.myMap.lat = parseFloat(value);
		});

		let colorValue = mapPane.querySelector('#color-cell').parentNode;
		this.pickColor({ parent: colorValue, cell: colorValue.querySelector('#color-cell') }, (color) => {

			colorValue.querySelector('#color-cell').value = color;
			let hexColor = ColorPicker.rgbToHex(color);
			colorValue.querySelector('#color-cell').setAttribute('value', hexColor); //set the value of the eventColor cell in the pane to the color
			this.sharePoint.properties.pane.content[this.key].settings.myMap.color = hexColor;
		});

		mapPane.querySelector('#longitude-cell').onChanged(value => {
			this.sharePoint.properties.pane.content[this.key].settings.myMap.lng = parseFloat(value);
		});

		mapPane.querySelector('#zoom-cell').onChanged(value => {
			this.sharePoint.properties.pane.content[this.key].settings.myMap.zoom = parseInt(value);
		});

		let markerValue = mapPane.querySelector('#marker-cell');
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

		mapPane.querySelector('#width-cell').onChanged(value => {
			draftDom.querySelector('#crater-map-div').css({ width: value });
			this.sharePoint.properties.pane.content[this.key].settings.myMap.width = value;
		});

		mapPane.querySelector('#height-cell').onChanged(value => {
			draftDom.querySelector('#crater-map-div').css({ height: value });
			this.sharePoint.properties.pane.content[this.key].settings.myMap.height = value;
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.querySelector('#crater-map-div').innerHTML = '';
			this.element.removeChild(this.element.querySelector('script'));
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
					let instaDiv = this.sharePoint.app.querySelector('.crater-instagram');
					instaDiv.removeChild(instaDiv.querySelector('.crater-instagram-content'));
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
				let errorMessage = this.sharePoint.app.querySelector('.crater-instagram-error');
				errorMessage.style.display = 'block';
			});
	}

	public renderInstagramPost(params?) {
		this.element = this.sharePoint.app.querySelector('.crater-instagram');
		let errorMessage = this.sharePoint.app.querySelector('.crater-instagram-error');
		errorMessage.style.display = 'none';
		this.key = this.element.dataset['key'];
		let display = params;
		let instaContent = this.element.querySelector('.crater-instagram-content');
		instaContent.innerHTML = '';

		let embedScript = document.createElement('script');
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
								element: 'input', name: 'width', attributes: { placeholder: 'Please enter a width' }, value: this.element.querySelector('.crater-instagram-content').css()['height'] || ''
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

		this.paneContent = this.sharePoint.app.querySelector('.crater-property-content').monitor();

		let instagramPane = this.paneContent.querySelector('.instagram-pane');
		let postUrl = instagramPane.querySelector('#postUrl-cell');
		postUrl.onChanged(value => {
			this.sharePoint.properties.pane.content[this.key].settings.myInstagram.instaURL = value;
			this.defaultURL = this.sharePoint.properties.pane.content[this.key].settings.myInstagram.instaURL;
			let finalWidth = this.sharePoint.properties.pane.content[this.key].settings.myInstagram.instaWidth || '&amp;minwidth=320&amp;maxwidth=320';
			const finalHide = (this.sharePoint.properties.pane.content[this.key].settings.myInstagram.instaCaption) ? '&amp;hidecaption=true' : '';

			this.sharePoint.properties.pane.content[this.key].draft.newEndPoint = this.sharePoint.properties.pane.content[this.key].settings.myInstagram.instaURL + `&amp;omitscript=true${finalHide}${finalWidth}`;
		});

		let hideCaption = instagramPane.querySelector('#hideCaption-cell');
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

		let changeWidth = instagramPane.querySelector('#width-cell');
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

		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
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

		let youtubeContent = youtube.querySelector('.crater-youtube-contents');
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

		this.paneContent = this.sharePoint.app.querySelector('.crater-property-content').monitor();

		let youtubePane = this.paneContent.querySelector('.youtube-pane');

		youtubePane.querySelector('#videoURL-cell').onChanged(value => {
			if (value.indexOf('.be/') !== -1) {
				youtubePane.querySelector('#videoURL').style.color = 'green';
				youtubePane.querySelector('#videoURL').textContent = 'Valid URL';
				let afterEmbed = value.split('.be/')[1];
				let newValue = 'https://www.youtube.com/embed/' + afterEmbed;
				this.sharePoint.properties.pane.content[this.key].settings.myYoutube.defaultVideo = newValue;
			} else {
				youtubePane.querySelector('#videoURL').style.color = 'red';
				youtubePane.querySelector('#videoURL').textContent = 'Invalid Video URL. Please right click on the video to get the video URL';
			}

		});

		youtubePane.querySelector('#width-cell').onChanged(value => {
			this.sharePoint.properties.pane.content[this.key].settings.myYoutube.width = value;
		});

		youtubePane.querySelector('#height-cell').onChanged(value => {
			this.sharePoint.properties.pane.content[this.key].settings.myYoutube.height = value;
		});

		this.paneContent.addEventListener('mutated', event => {
			// this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {

			let draftDomIframe = draftDom.querySelector('.crater-iframe');

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
			let crater = this.element.querySelector('.crater-facebook-content');
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
									this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.fb-page ').dataset.href
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
		this.paneContent = this.sharePoint.app.querySelector(".crater-property-content").monitor();

		let titlePane = this.paneContent.querySelector('.title-pane');
		let sizePane = this.paneContent.querySelector('.size-pane');
		let facebook = this.sharePoint.properties.pane.content[this.key].settings.facebook;


		titlePane.querySelector('#pageUrl-cell').onChanged(value => {
			facebook.url = value;
		});
		titlePane.querySelector('#Tabs-cell').onChanged(value => {
			facebook.tabs = value;
		});
		let coverCell = titlePane.querySelector('#hide-cover-cell');
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

		let hideFacePileCell = titlePane.querySelector('#hide-facepile-cell');
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

		let showHeaderCell = titlePane.querySelector('#small-header-cell');
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

		sizePane.querySelector('#width-cell').onChanged(value => {
			facebook.width = value;
		});

		sizePane.querySelector('#height-cell').onChanged(value => {
			facebook.height = value;
		});

		let adaptCell = sizePane.querySelector('#container-width-cell');
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

		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {

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
		const slider = this.element.querySelector('.crater-beforeAfter-contents').querySelector('.crater-handle');
		let isDown = false;
		let resizeDiv = this.element.querySelector('.crater-beforeAfter-contents').querySelector('.crater-after');
		let containerWidth = this.element.querySelector('.crater-beforeAfter-contents').offsetWidth + 'px';
		this.element.querySelector('.crater-after img').css({ "width": containerWidth });

		this.drags(slider, resizeDiv, this.element.querySelector('.crater-beforeAfter-contents'));
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

				let draggable = container.querySelector('.crater-draggable');

				if (!func.isnull(draggable)) draggable.css({ 'left': widthValue });
				container.addEventListener('mouseup', function () {
					this.classList.remove('crater-draggable');
					resizeElement.classList.remove('crater-resizable');
				});

				let resizable = container.querySelector('.crater-resizable');
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
									src: this.element.querySelector('.crater-beforeAfter-contents').querySelector('.crater-beforeImage').src
								}
							}),
							this.elementModifier.cell({
								element: 'img',
								name: 'after',
								dataAttributes: {
									style: { width: '400px', height: '400px' },
									src: this.element.querySelector('.crater-beforeAfter-contents').querySelector('.crater-after').querySelector('.crater-afterImage').src
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
		this.paneContent = this.sharePoint.app.querySelector('.crater-property-content').monitor();

		let afterCell = this.paneContent.querySelector('#after-cell').parentNode;
		this.uploadImage({ parent: afterCell }, (image) => {
			afterCell.querySelector('#after-cell').src = image.src;
			this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.beforeAfter-contents').querySelector('.crater-after').querySelector('.afterImage').src = image.src;
		});

		let beforeCell = this.paneContent.querySelector('#before-cell').parentNode;
		this.uploadImage({ parent: beforeCell }, (image) => {
			beforeCell.querySelector('#before-cell').src = image.src;
			this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.beforeAfter-contents').querySelector('.beforeImage').src = image.src;
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
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

		let content = event.querySelector(`.crater-event-content`);
		let locationElement = content.querySelectorAll('.crater-event-content-task-location') as any;

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
			let eventList = this.sharePoint.properties.pane.content[key].draft.dom.querySelector('.crater-event-content');
			let dateListRows = eventList.querySelectorAll('.crater-event-content-item');
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
								element: 'img', name: 'icon', attributes: {}, dataAttributes: { class: 'crater-icon', src: this.element.querySelector('.crater-event-title-imgIcon').src }
							}),
							this.elementModifier.cell({
								element: 'input', name: 'title', value: this.element.querySelector('.crater-event-title').textContent
							}),
							this.elementModifier.cell({
								element: 'input', name: 'backgroundcolor', value: this.element.querySelector('.crater-event-title').css()['background-color']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'color', value: this.element.querySelector('.crater-event-title').css().color
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontsize', value: this.element.querySelector('.crater-event-title').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'height', value: this.element.querySelector('.crater-event-title').css()['height'] || ''
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
								element: 'input', name: 'iconWidth', value: this.element.querySelector('.crater-event-content-item-icon-image').css()['width']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'iconHeight', value: this.element.querySelector('.crater-event-content-item-icon-image').css()['height']
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
								element: 'input', name: 'fontSize', value: this.element.querySelector('.crater-event-content-task-caption').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontFamily', value: this.element.querySelector('.crater-event-content-task-caption').css()['font-family']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'eventColor', value: this.element.querySelector('.crater-event-content-task-caption').css()['color']
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
								element: 'input', name: 'fontSize', value: this.element.querySelector('.crater-event-content-task-location-place').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontFamily', value: this.element.querySelector('.crater-event-content-task-location-place').css()['font-family']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'locationColor', value: this.element.querySelector('.crater-event-content-task-location-place').css().color
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
								element: 'input', name: 'fontSize', value: this.element.querySelector('.crater-event-content-task-location-duration').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontFamily', value: this.element.querySelector('.crater-event-content-task-location-duration').css()['font-family']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'durationColor', value: this.element.querySelector('.crater-event-content-task-location-duration').css()['color']
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
								element: 'input', name: 'daySize', value: this.element.querySelector('.crater-event-content-item-date-day').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'monthSize', value: this.element.querySelector('.crater-event-content-item-date-month').css()['font-size']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'fontFamily', value: this.element.querySelector('.crater-event-content-item-date').css()['font-family']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'dateColor', value: this.element.querySelector('.crater-event-content-item-date').css()['color']
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
		let cTime = strip(this.element.querySelector('#startTime').textContent)[0];
		let dTime = this.element.querySelector('#endTime').textContent;
		// let cDay = this.element.querySelector('.crater-event-content-item-date-day').textContent;
		// let cMonth = this.element.querySelector('.crater-event-content-item-date-month').textContent;
		// let gDate = new Date(`${cMonth} ${cDay}, 2019`);


		for (let i = 0; i < params.list.length; i++) {
			eventListPane.makeElement({
				element: 'div',
				attributes: { class: 'crater-event-item-pane row' },
				children: [
					this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-event-content-item' }),
					this.elementModifier.cell({
						element: 'img', name: 'icon', attributes: {}, dataAttributes: { class: 'crater-icon', src: params.list[i].querySelector('#icon').src }
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Task', attributes: { class: 'taskValue' }, value: params.list[i].querySelector('#eventTask').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Location', attributes: { class: 'locationValue' }, value: params.list[i].querySelector('#location').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Day', attribute: { class: 'crater-date dateValue' }, value: params.list[i].querySelector('#Day').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Month', attribute: { class: 'crater-date dateValue' }, value: params.list[i].querySelector('#Month').textContent
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
		this.paneContent = this.sharePoint.app.querySelector('.crater-property-content').monitor();
		//get the content and all the events
		let eventList = this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-event-content');
		let eventListRow = eventList.querySelectorAll('.crater-event-content-item');

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
							{ element: 'div', attributes: { class: 'crater-event-content-task-caption', id: 'eventTask' }, text: this.paneContent.querySelector(`.taskValue input`).value },
							{
								element: 'div', attributes: { class: 'crater-event-content-task-location' }, children: [
									{ element: 'img', attributes: { src: 'https://img.icons8.com/small/16/000000/clock.png' } },
									{ element: 'span', attributes: { class: 'crater-event-content-task-location-duration startTime' }, text: '' },
									{ element: 'span', attributes: { class: 'crater-event-content-task-location-duration endTime' }, text: '' },
									{ element: 'img', attributes: { src: 'https://img.icons8.com/small/16/000000/previous--location.png' } },
									{ element: 'span', attributes: { class: 'crater-event-content-task-location-place' }, text: this.paneContent.querySelector(`.locationValue input`).value }
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
				eventRowPane.querySelector('.crater-content-options').css({ visibility: 'visible' });
			});

			eventRowPane.addEventListener('mouseout', event => {
				eventRowPane.querySelector('.crater-content-options').css({ visibility: 'hidden' });
			});

			let iconCellParent = eventRowPane.querySelector('#icon-cell').parentNode;
			this.uploadImage({ parent: iconCellParent }, (image) => {
				iconCellParent.querySelector('#icon-cell').src = image.src;
				eventRowDom.querySelector('.crater-event-content-item-icon-image').src = image.src;
			});

			// get the values of the newly created row on the property - pane
			eventRowPane.querySelector('#Task-cell').onChanged(value => {
				eventRowDom.querySelector('.crater-event-content-task-caption').innerHTML = value;
			});

			eventRowPane.querySelector('#Location-cell').onChanged(value => {
				eventRowDom.querySelector('.stateCountry').innerHTML = value;
			});

			eventRowPane.querySelector('#Day-cell').onChanged(value => {
				eventRowDom.querySelector('.crater-event-content-item-date-day').innerHTML = value;
			});

			eventRowPane.querySelector('#Month-cell').onChanged(value => {
				eventRowDom.querySelector('.crater-event-content-item-date-month').innerHTML = value;
			});

			eventRowPane.querySelector('#start-cell').onChanged(value => {
				eventRowDom.querySelector('.startTime').innerHTML = value + ` - `;
			});

			eventRowPane.querySelector('#end-cell').onChanged(value => {
				eventRowDom.querySelector('.endTime').innerHTML = value;
			});

			eventRowPane.querySelector('.delete-crater-event-content-item').addEventListener('click', event => {
				eventRowDom.remove();
				eventRowPane.remove();
			});

			eventRowPane.querySelector('.add-before-crater-event-content-item').addEventListener('click', event => {
				let newEventRowDom = eventListRowDomPrototype.cloneNode(true);
				let neweventRowPane = eventListRowPanePrototype.cloneNode(true);

				eventRowDom.before(newEventRowDom);
				eventRowPane.before(neweventRowPane);
				eventRowHandler(neweventRowPane, newEventRowDom);
			});

			eventRowPane.querySelector('.add-after-crater-event-content-item').addEventListener('click', event => {
				let newEventRowDom = eventListRowDomPrototype.cloneNode(true);
				let newEventRowPane = eventListRowPanePrototype.cloneNode(true);

				eventRowDom.after(newEventRowDom);
				eventRowPane.after(newEventRowPane);

				eventRowHandler(newEventRowPane, newEventRowDom);
			});
		};

		let draftDom = this.sharePoint.properties.pane.content[this.key].draft.dom;

		let titlePane = this.paneContent.querySelector('.title-pane');
		let eventIconRowPane = this.paneContent.querySelector('.event-icon-row-pane');
		let eventTitleRowPane = this.paneContent.querySelector('.event-title-row-pane');
		let eventLocationRowPane = this.paneContent.querySelector('.event-location-row-pane');
		let eventDurationRowPane = this.paneContent.querySelector('.event-duration-row-pane');
		let eventDateRowPane = this.paneContent.querySelector('.event-date-row-pane');

		let iconParent = titlePane.querySelector('#icon-cell').parentNode;

		let eventColorParent = eventTitleRowPane.querySelector('#eventColor-cell').parentNode;
		this.pickColor({ parent: eventColorParent, cell: eventColorParent.querySelector('#eventColor-cell') }, (color) => {
			draftDom.querySelectorAll('.crater-event-content-task-caption').forEach(element => {
				element.css({ color });//get the color of the event font in the draftDom
			});
			eventColorParent.querySelector('#eventColor-cell').value = color;
			eventColorParent.querySelector('#eventColor-cell').setAttribute('value', color); //set the value of the eventColor cell in the pane to the color
		});

		let dateColorParent = eventDateRowPane.querySelector('#dateColor-cell').parentNode;
		this.pickColor({ parent: dateColorParent, cell: dateColorParent.querySelector('#dateColor-cell') }, (color) => {
			draftDom.querySelectorAll('.crater-event-content-item-date').forEach(element => {
				element.css({ color });//get the color of the event font in the draftDom
			});
			dateColorParent.querySelector('#dateColor-cell').value = color;
			dateColorParent.querySelector('#dateColor-cell').setAttribute('value', color); //set the value of the eventColor cell in the pane to the color
		});

		let locationColorParent = eventLocationRowPane.querySelector('#locationColor-cell').parentNode;
		this.pickColor({ parent: locationColorParent, cell: locationColorParent.querySelector('#locationColor-cell') }, (color) => {
			draftDom.querySelectorAll('.crater-event-content-task-location-place').forEach(element => {
				element.css({ color });//get the color of the event font in the draftDom
			});
			locationColorParent.querySelector('#locationColor-cell').value = color;
			locationColorParent.querySelector('#locationColor-cell').setAttribute('value', color); //set the value of the eventColor cell in the pane to the color
		});

		let durationColorParent = eventDurationRowPane.querySelector('#durationColor-cell').parentNode;
		this.pickColor({ parent: durationColorParent, cell: durationColorParent.querySelector('#durationColor-cell') }, (color) => {
			draftDom.querySelectorAll('.crater-event-content-task-location-duration').forEach(element => {
				element.css({ color });//get the color of the event font in the draftDom
			});
			durationColorParent.querySelector('#durationColor-cell').value = color;
			durationColorParent.querySelector('#durationColor-cell').setAttribute('value', color); //set the value of the eventColor cell in the pane to the color
		});
		this.uploadImage({ parent: iconParent }, (image) => {
			iconParent.querySelector('#icon-cell').src = image.src;
			draftDom.querySelector('.crater-event-title-imgIcon').src = image.src;
		});
		titlePane.querySelector('#title-cell').onChanged(value => {
			draftDom.querySelector('.crater-event-title-captionTitle').innerHTML = value;
		});

		let backgroundColorCell = titlePane.querySelector('#backgroundcolor-cell').parentNode;
		this.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.querySelector('#backgroundcolor-cell') }, (backgroundColor) => {
			draftDom.querySelector('.crater-event-title').css({ backgroundColor });
			backgroundColorCell.querySelector('#backgroundcolor-cell').value = backgroundColor;
			backgroundColorCell.querySelector('#backgroundcolor-cell').setAttribute('value', backgroundColor);
		});

		let colorCell = titlePane.querySelector('#color-cell').parentNode;
		this.pickColor({ parent: colorCell, cell: colorCell.querySelector('#color-cell') }, (color) => {
			draftDom.querySelector('.crater-event-title').css({ color });
			colorCell.querySelector('#color-cell').value = color;
			colorCell.querySelector('#color-cell').setAttribute('value', color);
		});


		titlePane.querySelector('#fontsize-cell').onChanged(value => {
			draftDom.querySelector('.crater-event-title').css({ fontSize: value });
		});

		titlePane.querySelector('#height-cell').onChanged(value => {
			draftDom.querySelector('.crater-event-title').css({ height: value });
		});

		eventIconRowPane.querySelector('#iconWidth-cell').onChanged(value => {
			draftDom.querySelectorAll('.crater-event-content-item-icon-image').forEach(element => {
				element.css({ width: value });
			});
		});
		eventIconRowPane.querySelector('#iconHeight-cell').onChanged(value => {
			draftDom.querySelectorAll('.crater-event-content-item-icon-image').forEach(element => {
				element.css({ height: value });
			});
		});

		eventTitleRowPane.querySelector('#fontSize-cell').onChanged(value => {
			draftDom.querySelectorAll('.crater-event-content-task-caption').forEach(element => {
				element.css({ fontSize: value });
			});
		});
		eventTitleRowPane.querySelector('#fontFamily-cell').onChanged(value => {
			draftDom.querySelectorAll('.crater-event-content-task-caption').forEach(element => {
				element.css({ fontFamily: value });
			});
		});

		eventLocationRowPane.querySelector('#fontSize-cell').onChanged(value => {
			draftDom.querySelectorAll('.crater-event-content-task-location-place').forEach(element => {
				element.css({ fontSize: value });
			});
		});
		eventLocationRowPane.querySelector('#fontFamily-cell').onChanged(value => {
			draftDom.querySelectorAll('.crater-event-content-task-location-place').forEach(element => {
				element.css({ fontFamily: value });
			});
		});

		eventDurationRowPane.querySelector('#fontSize-cell').onChanged(value => {
			draftDom.querySelectorAll('.crater-event-content-task-location-duration').forEach(element => {
				element.css({ fontSize: value });
			});
		});
		eventDurationRowPane.querySelector('#fontFamily-cell').onChanged(value => {
			draftDom.querySelectorAll('.crater-event-content-task-location-duration').forEach(element => {
				element.css({ fontFamily: value });
			});
		});

		eventDateRowPane.querySelector('#daySize-cell').onChanged(value => {
			draftDom.querySelectorAll('.crater-event-content-item-date-day').forEach(element => {
				element.css({ fontSize: value });
			});
		});
		eventDateRowPane.querySelector('#monthSize-cell').onChanged(value => {
			draftDom.querySelectorAll('.crater-event-content-item-date-month').forEach(element => {
				element.css({ fontSize: value });
			});
		});

		eventDateRowPane.querySelector('#fontFamily-cell').onChanged(value => {
			draftDom.querySelectorAll('.crater-event-content-item-date').forEach(element => {
				element.css({ fontFamily: value });
			});
		});
		//appends the dom and pane prototypes to the dom and pane when you click add new
		this.paneContent.querySelector('.new-component').addEventListener('click', event => {
			let newEventRowDom = eventListRowDomPrototype.cloneNode(true);
			let newEventRowPane = eventListRowPanePrototype.cloneNode(true);

			eventList.append(newEventRowDom);//c
			this.paneContent.querySelector('.list-pane').append(newEventRowPane);
			eventRowHandler(newEventRowPane, newEventRowDom);
		});

		let paneItems = this.paneContent.querySelectorAll('.crater-event-item-pane');
		paneItems.forEach((eventRow, position) => {
			eventRowHandler(eventRow, eventListRow[position]);
		});

		let showHeader = titlePane.querySelector('#toggleTitle-cell');
		showHeader.addEventListener('change', e => {

			switch (showHeader.value.toLowerCase()) {
				case "hide":
					draftDom.querySelectorAll('.crater-event-title').forEach(element => {
						element.style.display = "none";
					});
					break;
				case "show":
					draftDom.querySelectorAll('.crater-event-title').forEach(element => {
						element.style.display = "flex";
					});
					break;
				default:
					draftDom.querySelectorAll('.crater-event-title').forEach(element => {
						element.style.display = "none";
					});
			}
		});

		//to hide or show properties
		let showIcon = eventIconRowPane.querySelector('#toggleIcon-cell');

		showIcon.addEventListener('change', e => {

			switch (showIcon.value.toLowerCase()) {
				case "hide":
					draftDom.querySelectorAll('.crater-event-content-item-icon-image').forEach(element => {
						element.style.display = "none";
					});
					break;
				case "show":
					draftDom.querySelectorAll('.crater-event-content-item-icon-image').forEach(element => {
						element.style.display = "block";
					});
					break;
				default:
					draftDom.querySelectorAll('.crater-event-content-item-icon-image').forEach(element => {
						element.style.display = "none";
					});
			}
		});

		let showTitle = eventTitleRowPane.querySelector('#toggleTitle-cell');
		showTitle.addEventListener('change', e => {

			switch (showTitle.value.toLowerCase()) {
				case "hide":
					draftDom.querySelectorAll('.crater-event-content-task-caption').forEach(element => {
						element.style.visibility = "hidden";
					});
					break;
				case "show":
					draftDom.querySelectorAll('.crater-event-content-task-caption').forEach(element => {
						element.style.visibility = "visible";
					});
					break;
				default:
					draftDom.querySelectorAll('.crater-event-content-task-caption').forEach(element => {
						element.style.visibility = "hidden";
					});
			}
		});

		let showLocation = eventLocationRowPane.querySelector('#toggleLocation-cell');
		showLocation.addEventListener('change', e => {

			switch (showLocation.value.toLowerCase()) {
				case "hide":
					draftDom.querySelectorAll('.crater-event-content-task-location-place').forEach(element => {
						element.style.visibility = "hidden";
						element.previousSibling.style.visibility = "hidden";
					});
					break;
				case "show":
					draftDom.querySelectorAll('.crater-event-content-task-location-place').forEach(element => {
						element.style.visibility = "visible";
						element.previousSibling.style.visibility = "visible";
					});
					break;
				default:
					draftDom.querySelectorAll('.crater-event-content-task-location-place').forEach(element => {
						element.style.visibility = "hidden";
						element.previousSibling.style.visibility = "hidden";
					});
			}
		});

		let showDuration = eventDurationRowPane.querySelector('#toggleDuration-cell');
		showDuration.addEventListener('change', e => {

			switch (showDuration.value.toLowerCase()) {
				case "hide":
					draftDom.querySelectorAll('.crater-event-content-task-location-duration').forEach(element => {
						element.style.visibility = "hidden";
						element.previousSibling.style.visibility = "hidden";
					});
					break;
				case "show":
					draftDom.querySelectorAll('.crater-event-content-task-location-duration').forEach(element => {
						element.style.visibility = "visible";
						element.previousSibling.style.visibility = "visible";
					});
					break;
				default:
					draftDom.querySelectorAll('.crater-event-content-task-location-duration').forEach(element => {
						element.style.visibility = "hidden";
						element.previousSibling.style.visibility = "hidden";
					});
			}
		});

		let showDate = eventDateRowPane.querySelector('#toggleDate-cell');
		showDate.addEventListener('change', e => {

			switch (showDate.value.toLowerCase()) {
				case "hide":
					draftDom.querySelectorAll('.crater-event-content-item-date').forEach(element => {
						element.style.display = "none";
					});
					break;
				case "show":
					draftDom.querySelectorAll('.crater-event-content-item-date').forEach(element => {
						element.style.display = "block";
					});
					break;
				default:
					draftDom.querySelectorAll('.crater-event-content-item-date').forEach(element => {
						element.style.display = "none";
					});
			}
		});

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
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

		let paneConnection = this.sharePoint.app.querySelector('.crater-property-connection');

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
		updateWindow.querySelector('#update-element').addEventListener('click', event => {
			event.preventDefault();
			data.title = updateWindow.querySelector('#meta-data-title').value;
			data.icon = updateWindow.querySelector('#meta-data-icon').value;
			data.day = updateWindow.querySelector('#meta-data-day').value;
			data.month = updateWindow.querySelector('#meta-data-month').value;
			data.location = updateWindow.querySelector('#meta-data-location').value;
			data.start = updateWindow.querySelector('#meta-data-start').value;
			data.end = updateWindow.querySelector('#meta-data-end').value;
			source = func.extractFromJsonArray(data, params.source);

			let newContent = this.render({ source });
			draftDom.querySelector('.crater-event-content').innerHTML = newContent.querySelector('.crater-event-content').innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = draftDom.outerHTML;
			this.paneContent.querySelector('.list-pane').innerHTML = this.generatePaneContent({ list: newContent.querySelectorAll('.crater-event-content-item') }).innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
		});


		if (!func.isnull(paneConnection)) {
			paneConnection.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
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

		power.querySelector('.login_form').innerHTML += form;

		this.key = power.dataset.key;
		this.sharePoint.properties.pane.content[this.key].settings.myPowerBi = { showNavContent: '', showFilter: '', loginType: '', code: '', username: '', tenantID: "90fa49f0-0a6c-4170-abed-92ed96ba67ca", clientSecret: 'FUq.Y0@BN4byWh6B8.H:et:?F/VX2-3a', password: '', clientId: '9605a407-7c23-4dc8-bd90-997fbc254d38', accessToken: '', embedToken: '', embedUrl: '', reports: [], reportName: [], groups: [], groupName: [], dashboards: [], dashBoardName: [], tiles: [], tileName: [], width: '100%', height: '300px' };
		window.onerror = (msg, url, lineNumber, columnNumber, error) => {
			console.log(msg, url, lineNumber, columnNumber, error);
		};
		power.querySelector('#master').addEventListener('click', event => {
			let loginForm = power.querySelector('.login_form') as any;
			let powerConnect = power.querySelector('.crater-power-connect') as any;

			powerConnect.style.display = "none";
			loginForm.style.display = 'block';
			let username = power.querySelector('#power-username') as any;
			let password = power.querySelector('#power-password') as any;
			let errorDisplay = power.querySelector('#emptyField') as any;
			let loginButton = power.querySelector('#login-submit') as any;
			let cancelButton = power.querySelector('#cancelbtn') as any;
			let renderBox = power.querySelector('#renderContainer') as any;

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
							renderBox.querySelector('#render-error').style.display = 'none';
							renderBox.style.display = "block";
							this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.loginType = 'master';
							loginForm.style.display = "none";
							if (renderBox.querySelector('.connected')) renderBox.querySelector('.connected').remove();
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
							renderBox.querySelector('#master').addEventListener('click', ev => {
								loginButton.innerHTML = 'Login';
								if (power.querySelector('.login_form').style.display = 'none') power.querySelector('.login_form').style.display = 'block';
							});


						}
					});
				}
			});

			cancelButton.addEventListener('click', e => {
				loginForm.style.display = 'none';
				if (!renderBox.querySelector('.connected')) powerConnect.style.display = 'grid';
			});

		});
		power.querySelector('#normal').addEventListener('click', event => {
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
							let powerConnect = power.querySelector('.crater-power-connect') as any;
							powerConnect.style.display = "none";
							let renderBox = power.querySelector('#renderContainer') as any;
							renderBox.querySelector('#render-error').style.display = 'none';
							renderBox.style.display = 'block';
							power.querySelector('.crater-power-timer').style.display = "none";
							if (renderBox.querySelector('.connected')) renderBox.querySelector('.connected').remove();
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
		let loadingButton = document.querySelector('#login-submit') as any;
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
		document.querySelector('.crater-power-timer').style.display = 'block';
		clearInterval(1);
		setInterval(() => {
			if (this.counter === 0) {
				//@ts-ignore
				document.querySelector('.crater-power-counter').render({
					element: 'img', attributes: { class: 'crater-icon', src: this.sharePoint.images.loading, style: { width: '20px', height: '20px' } }
				});
			}
			else {
				this.counter--;
				//@ts-ignore
				document.querySelector('.crater-power-counter').textContent = 'Please wait ' + this.counter + ' Seconds';
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
			document.querySelector('#render-error').style.display = 'block';
			document.querySelector('#render-text').textContent = error.message;
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
			document.querySelector('#render-error').style.display = 'block';
			document.querySelector('#render-text').textContent = error.message;
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
			document.querySelector('#render-error').style.display = 'block';
			document.querySelector('#render-text').textContent = error.message;
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
			document.querySelector('#render-error').style.display = 'block';
			document.querySelector('#render-text').textContent = error.message;
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
			document.querySelector('#renderContainer').style.display = "block";
			document.querySelector('#render-text').textContent = error.message;

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
			document.querySelector('#renderContainer').style.display = "block";
			document.querySelector('#render-text').textContent = error.message;
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
			document.querySelector('#renderContainer').style.display = "block";
			document.querySelector('#render-text').textContent = error.message;
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
			document.querySelector('#render-error').style.display = 'block';
			document.querySelector('#render-text').textContent = error.message;
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
			document.querySelector('#render-error').style.display = 'block';
			document.querySelector('#render-text').textContent = error.message;
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
			document.querySelector('#render-error').style.display = 'block';
			document.querySelector('#render-text').textContent = error.message;
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
			document.querySelector('#render-error').style.display = 'block';
			document.querySelector('#render-text').textContent = error.message;
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
			document.querySelector('#render-error').style.display = 'block';
			document.querySelector('#render-text').textContent = error.message;
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

					if (document.querySelector('#getWork')) {
						let errorText = document.querySelector('#render-text') as any;
						errorText.textContent = '';
						//@ts-ignore
						document.querySelector('#renderContainer').style.display = "none";
						document.querySelector('#getWork').remove();
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
			let powDiv = document.querySelector('#render-error') as any;
			//@ts-ignore
			document.querySelector('#renderContainer').style.display = 'block';
			powDiv.querySelector('#render-text').textContent = `Couldn't Fetch Workspaces!`;
			powDiv.makeElement({
				element: 'div', children: [
					{ element: 'button', attributes: { id: 'getWork', class: 'user-button' }, text: 'Retry' }
				]
			});
			powDiv.querySelector('#getWork').addEventListener('click', even => {
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
			document.querySelector('#login-submit').style.zIndex = 0;
			document.querySelector('#login-submit').innerHTML = 'Login';
			document.querySelector('.crater-power-counter').innerHTML = "Sorry, there was an error. Please, click the connect button to try again";
			//@ts-ignore
			document.querySelector('#renderContainer').style.display = "block";
			document.querySelector('#render-text').textContent = `Please make sure your details are valid!`;
		});

		return promise;
		// -------------------------------------------
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
			let embedContainer = document.querySelector('#renderContainer') as any;
			//@ts-ignore
			let reportRefresh = powerbi.get(embedContainer);

			reportRefresh.setAccessToken(response).then(() => {
				this.tokenListener({ tokenExpiration: draftPower.expiration, minutesToRefresh: 2 });
			});
		}).catch(error => {
			//@ts-ignore
			document.querySelector('#renderContainer').style.display = "block";
			document.querySelector('#render-text').textContent = error.message;
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
		let self = this;
		let promise = new Promise((res, rej) => {
			let result;
			let request = new XMLHttpRequest();
			request.onreadystatechange = function (e) {
				if (this.readyState == 4 && this.status == 200) {
					result = request.responseText;
					console.log('getting embed token...');
					draftPower.embedToken = JSON.parse(result).token;
					draftPower.expiration = JSON.parse(result).expiration;

					if (self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.changed) {
						if (document.querySelector('#renderContainer')) {
							document.querySelector('#renderContainer').remove();
							let powerContainer = document.querySelector('.crater-power-container') as any;
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
			document.querySelector('#render-error').style.display = 'block';
			document.querySelector('#render-text').textContent = error.message;
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
			const reportContainer = document.querySelector('#renderContainer') as any;

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
			document.querySelector('.login_form').style.display = "block";
			//@ts-ignore
			document.querySelector('#render-error').style.display = 'block';
			document.querySelector('#render-text').textContent = error.message;
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
				this.element.querySelector('.crater-power-container').css({ width: draftPower.width, height: draftPower.height });
				this.element.querySelector('#renderContainer').css({ width: '100%', height: draftPower.height });
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
				let userList = this.sharePoint.properties.pane.content[key].draft.dom.querySelector('.crater-power-container');
				let userListRows = userList.querySelectorAll('.user');

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
									element: 'input', name: 'backgroundcolor', value: this.element.querySelector('.user').css()['background-color']
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
									element: 'input', name: 'fontSize', value: this.element.querySelector('.user-button').css()['font-size']
								}),
								this.elementModifier.cell({
									element: 'input', name: 'backgroundcolor', value: this.element.querySelector('.user-button').css()['background-color']
								}),
								this.elementModifier.cell({
									element: 'input', name: 'fontFamily', value: this.element.querySelector('.user-button').css()['font-family']
								}),
								this.elementModifier.cell({
									element: 'input', name: 'color', value: this.element.querySelector('.user-button').css()['color']
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
									element: 'input', name: 'fontSize', value: this.element.querySelector('.user-recommended').css()['font-size']
								}),
								this.elementModifier.cell({
									element: 'input', name: 'fontFamily', value: this.element.querySelector('.user-recommended').css()['font-family']
								}),
								this.elementModifier.cell({
									element: 'input', name: 'color', value: this.element.querySelector('.user-recommended').css()['color']
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
									element: 'input', name: 'fontSize', value: this.element.querySelector('.power-text').css()['font-size']
								}),
								this.elementModifier.cell({
									element: 'input', name: 'fontFamily', value: this.element.querySelector('.power-text').css()['font-family']
								}),
								this.elementModifier.cell({
									element: 'input', name: 'color', value: this.element.querySelector('.power-text').css()['color']
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
									element: 'input', name: 'fontSize', value: this.element.querySelector('.user-header').css()['font-size']
								}),
								this.elementModifier.cell({
									element: 'input', name: 'fontFamily', value: this.element.querySelector('.user-header').css()['font-family']
								}),
								this.elementModifier.cell({
									element: 'input', name: 'color', value: this.element.querySelector('.user-header').css()['color']
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
						element: 'img', name: 'image', attributes: {}, dataAttributes: { class: 'crater-icon', src: params.list[i].querySelector('.crater-power-image').src }
					}),
					this.elementModifier.cell({
						element: 'input', name: 'header', attribute: { class: 'crater-user-header' }, value: params.list[i].querySelector('.user-header').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'info-text', value: params.list.title || params.list[i].querySelector('.power-text').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'recommended', value: params.list.subTitle || params.list[i].querySelector('.user-recommended').textContent
					}),
					this.elementModifier.cell({
						element: 'input', name: 'button', value: params.list.body || params.list[i].querySelector('.user-button').textContent
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
		this.paneContent = this.sharePoint.app.querySelector('.crater-property-content').monitor();

		window.onerror = (msg, url, lineNumber, columnNumber, error) => {
			console.log(msg, url, lineNumber, columnNumber, error);
		};

		if (this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.loginType.length !== 0) {

			let powerPane = this.paneContent.querySelector('.power-pane');
			let sizePane = this.paneContent.querySelector('.size-pane');
			if (this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.loginType.toLowerCase() === 'user') {
				powerPane.querySelector('#user-rights-cell').parentElement.remove();
			} else {
				powerPane.querySelector('#user-rights-cell').onChanged(value => {
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
					if (powerPane.querySelector('#page-cell')) {
						powerPane.querySelector('#page-cell').onChanged(value => {
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
					if (powerPane.querySelector('#tile-cell')) {
						powerPane.querySelector('#tile-cell').onChanged(value => {

							for (let tile = 0; tile < self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.tiles.length; tile++) {
								for (let property in self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.tiles[tile]) {
									if (value === self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.tiles[tile].tileName) {
										if (powerPane.querySelector('#filter-panel-cell')) powerPane.querySelector('#filter-panel-cell').parentElement.parentElement.remove();
										if (powerPane.querySelector('#navigation-cell')) powerPane.querySelector('#navigation-cell').parentElement.parentElement.remove();
										if (powerPane.querySelector('#page-cell')) powerPane.querySelector('#page-cell').parentElement.remove();

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
										if (powerPane.querySelector('#page-cell')) powerPane.querySelector('#page-cell').parentElement.remove();
										if (powerPane.querySelector('#filter-panel-cell')) powerPane.querySelector('#filter-panel-cell').parentElement.parentElement.remove();
										if (powerPane.querySelector('#navigation-cell')) powerPane.querySelector('#navigation-cell').parentElement.parentElement.remove();
									}
								}
							}

							if (value.toLowerCase() !== 'show full dashboard') {
								powerPane.querySelector('.power-embed').makeElement({
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

								powerPane.querySelector('#filter-panel-cell').onChanged(val => {
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

								powerPane.querySelector('#navigation-cell').onChanged(val => {
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
					if (powerPane.querySelector('#view-cell')) {
						powerPane.querySelector('#view-cell').onChanged(value => {
							if (powerPane.querySelector('#page-cell')) powerPane.querySelector('#page-cell').parentElement.remove();
							if (powerPane.querySelector('#filter-panel-cell')) powerPane.querySelector('#filter-panel-cell').parentElement.parentElement.remove();
							if (powerPane.querySelector('#navigation-cell')) powerPane.querySelector('#navigation-cell').parentElement.parentElement.remove();

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
							powerPane.querySelector('.power-embed').makeElement({
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

							powerPane.querySelector('#filter-panel-cell').onChanged(val => {
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

							powerPane.querySelector('#navigation-cell').onChanged(val => {
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
					if (powerPane.querySelector('#view-cell')) {
						powerPane.querySelector('#view-cell').onChanged(value => {
							if (powerPane.querySelector('#filter-panel-cell')) powerPane.querySelector('#filter-panel-cell').parentElement.parentElement.remove();
							if (powerPane.querySelector('#navigation-cell')) powerPane.querySelector('#navigation-cell').parentElement.parentElement.remove();
							if (powerPane.querySelector('#page-cell')) powerPane.querySelector('#page-cell').parentElement.remove();
							if (powerPane.querySelector('#tile-cell')) powerPane.querySelector('#tile-cell').parentElement.remove();
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
							powerPane.querySelector('.power-embed').makeElement({
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

			powerPane.querySelector('#WorkSpace-cell').onChanged(value => {
				this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.changed = true;
				if (powerPane.querySelector('#embed-type-cell')) powerPane.querySelector('#embed-type-cell').parentElement.remove();
				if (powerPane.querySelector('#view-cell')) powerPane.querySelector('#view-cell').parentElement.remove();
				if (powerPane.querySelector('#tile-cell')) powerPane.querySelector('#tile-cell').parentElement.remove();
				if (powerPane.querySelector('#page-cell')) powerPane.querySelector('#page-cell').parentElement.remove();
				if (powerPane.querySelector('#filter-panel-cell')) powerPane.querySelector('#filter-panel-cell').parentElement.parentElement.remove();
				if (powerPane.querySelector('#navigation-cell')) powerPane.querySelector('#navigation-cell').parentElement.parentElement.remove();
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

				powerPane.querySelector('.power-embed').makeElement({
					element: 'div', children: [
						self.elementModifier.cell({
							element: 'select', name: 'embed-type', options: ['Dashboard', 'Report']
						})
					]
				});

				powerPane.querySelector('#embed-type-cell').onChanged(val => {
					switch (val.toLowerCase()) {
						case 'report':
							if (powerPane.querySelector('#filter-panel-cell')) powerPane.querySelector('#filter-panel-cell').parentElement.parentElement.remove();
							if (powerPane.querySelector('#navigation-cell')) powerPane.querySelector('#navigation-cell').parentElement.parentElement.remove();
							if (powerPane.querySelector('#page-cell')) powerPane.querySelector('#page-cell').parentElement.remove();
							if (powerPane.querySelector('#view-cell')) powerPane.querySelector('#view-cell').parentElement.remove();
							if (powerPane.querySelector('#tile-cell')) powerPane.querySelector('#tile-cell').parentElement.remove();
							powerPane.querySelector('.power-embed').makeElement({
								element: 'div', children: [
									this.elementModifier.cell({
										element: 'select', name: 'view', options: this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.reportName
									})
								]
							});
							self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.tileId = '';
							self.sharePoint.properties.pane.content[self.key].settings.myPowerBi.dashboardId = '';
							// if (this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.reportName.indexOf('No Reports') !== -1) {
							// 	powerPane.querySelector('#view-cell').options[0].selected = true;
							// 	powerPane.querySelector('#view-cell').disabled = true;
							// }
							this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.embedType = val.toLowerCase();
							changedReport();
							break;
						case 'dashboard':
							if (powerPane.querySelector('#filter-panel-cell')) powerPane.querySelector('#filter-panel-cell').parentElement.parentElement.remove();
							if (powerPane.querySelector('#navigation-cell')) powerPane.querySelector('#navigation-cell').parentElement.parentElement.remove();
							if (powerPane.querySelector('#page-cell')) powerPane.querySelector('#page-cell').parentElement.remove();
							if (powerPane.querySelector('#view-cell')) powerPane.querySelector('#view-cell').parentElement.remove();
							if (powerPane.querySelector('#tile-cell')) powerPane.querySelector('#tile-cell').parentElement.remove();

							powerPane.querySelector('.power-embed').makeElement({
								element: 'div', children: [
									this.elementModifier.cell({
										element: 'select', name: 'view', options: this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.dashBoardName
									})
								]
							});
							// if (this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.dashBoardName.indexOf('No Dashboards') !== -1) {
							// 	powerPane.querySelector('#view-cell').options[0].selected = true;
							// 	powerPane.querySelector('#view-cell').disabled = true;
							// }
							this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.embedType = val.toLowerCase();
							changedDashboard();
							break;
					}
				});
			});

			sizePane.querySelector('#display-cell').onChanged(value => {
				switch (value) {
					case '16:9 (1280px x 720px)':
						if (sizePane.querySelector('#width-cell')) sizePane.querySelector('#width-cell').parentElement.remove();
						if (sizePane.querySelector('#height-cell')) sizePane.querySelector('#height-cell').parentElement.remove();
						this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.width = '1280px';
						this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.height = '720px';
						break;
					case '4:3 (1000px x 750px)':
						if (sizePane.querySelector('#width-cell')) sizePane.querySelector('#width-cell').parentElement.remove();
						if (sizePane.querySelector('#height-cell')) sizePane.querySelector('#height-cell').parentElement.remove();
						this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.width = '1000px';
						this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.height = '750px';
						break;
					case 'Custom Size':
						sizePane.querySelector('.power-embed').makeElement({
							element: 'div', children: [
								this.elementModifier.cell({
									element: 'input', name: 'width', value: this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.width
								}),
								this.elementModifier.cell({
									element: 'input', name: 'height', value: this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.height
								})
							]
						});
						sizePane.querySelector('#width-cell').onChanged(val => {
							this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.width = val;
						});
						sizePane.querySelector('#height-cell').onChanged(val => {
							this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.height = val;
						});
						break;
				}
			});
		}
		else {
			let layoutRowPane = this.paneContent.querySelector('.layout-pane');
			let layoutButtonRowPane = this.paneContent.querySelector('.layout-button-row-pane');
			let layoutRecommendedRowPane = this.paneContent.querySelector('.layout-recommended-row-pane');
			let layoutInfoRowPane = this.paneContent.querySelector('.layout-info-row-pane');
			let layoutHeaderRowPane = this.paneContent.querySelector('.layout-header-row-pane');
			let userList = this.element.querySelector('.crater-power-connect');
			let userListRow = userList.querySelectorAll('.user');

			let layoutBackgroundParent = layoutRowPane.querySelector('#backgroundcolor-cell').parentNode;
			this.pickColor({ parent: layoutBackgroundParent, cell: layoutBackgroundParent.querySelector('#backgroundcolor-cell') }, (backgroundColor) => {
				this.element.querySelectorAll('.login-container').forEach(element => {
					element.css({ backgroundColor });
				});
				layoutBackgroundParent.querySelector('#backgroundcolor-cell').value = backgroundColor;
				layoutBackgroundParent.querySelector('#backgroundcolor-cell').setAttribute('value', backgroundColor); //set the value of the eventColor cell in the pane to the color
			});

			let layoutButtonParent = layoutButtonRowPane.querySelector('#backgroundcolor-cell').parentNode;
			this.pickColor({ parent: layoutButtonParent, cell: layoutButtonParent.querySelector('#backgroundcolor-cell') }, (backgroundColor) => {
				this.element.querySelectorAll('.user-button').forEach(element => {
					element.css({ backgroundColor });
				});
				layoutButtonParent.querySelector('#backgroundcolor-cell').value = backgroundColor;
				layoutButtonParent.querySelector('#backgroundcolor-cell').setAttribute('value', backgroundColor); //set the value of the eventColor cell in the pane to the color
			});

			let layoutButtonColor = layoutButtonRowPane.querySelector('#color-cell').parentNode;
			this.pickColor({ parent: layoutButtonColor, cell: layoutButtonColor.querySelector('#color-cell') }, (color) => {
				this.element.querySelectorAll('.user-button').forEach(element => {
					element.css({ color });
				});
				layoutButtonColor.querySelector('#color-cell').value = color;
				layoutButtonColor.querySelector('#color-cell').setAttribute('value', color);
			});

			let layoutRecommendedColor = layoutRecommendedRowPane.querySelector('#color-cell').parentNode;
			this.pickColor({ parent: layoutRecommendedColor, cell: layoutRecommendedColor.querySelector('#color-cell') }, (color) => {
				this.element.querySelectorAll('.user-recommended').forEach(element => {
					element.css({ color });
				});
				layoutRecommendedColor.querySelector('#color-cell').value = color;
				layoutRecommendedColor.querySelector('#color-cell').setAttribute('value', color);
			});

			let layoutInfoColor = layoutInfoRowPane.querySelector('#color-cell').parentNode;
			this.pickColor({ parent: layoutInfoColor, cell: layoutInfoColor.querySelector('#color-cell') }, (color) => {
				this.element.querySelectorAll('.power-text').forEach(element => {
					element.css({ color });
				});
				layoutInfoColor.querySelector('#color-cell').value = color;
				layoutInfoColor.querySelector('#color-cell').setAttribute('value', color);
			});

			let layoutHeaderColor = layoutHeaderRowPane.querySelector('#color-cell').parentNode;
			this.pickColor({ parent: layoutHeaderColor, cell: layoutHeaderColor.querySelector('#color-cell') }, (color) => {
				this.element.querySelectorAll('.user-header').forEach(element => {
					element.css({ color });
				});
				layoutHeaderColor.querySelector('#color-cell').value = color;
				layoutHeaderColor.querySelector('#color-cell').setAttribute('value', color);
			});

			layoutButtonRowPane.querySelector('#fontSize-cell').onChanged(value => {
				this.element.querySelectorAll('.user-button').forEach(element => {
					element.css({ fontSize: value });
				});
			});

			layoutRecommendedRowPane.querySelector('#fontSize-cell').onChanged(value => {
				this.element.querySelectorAll('.user-recommended').forEach(element => {
					element.css({ fontSize: value });
				});
			});

			layoutInfoRowPane.querySelector('#fontSize-cell').onChanged(value => {
				this.element.querySelectorAll('.power-text').forEach(element => {
					element.css({ fontSize: value });
				});
			});

			layoutHeaderRowPane.querySelector('#fontSize-cell').onChanged(value => {
				this.element.querySelectorAll('.user-header').forEach(element => {
					element.css({ fontSize: value });
				});
			});

			layoutButtonRowPane.querySelector('#fontFamily-cell').onChanged(value => {
				this.element.querySelectorAll('.user-button').forEach(element => {
					element.css({ fontFamily: value });
				});
			});

			layoutRecommendedRowPane.querySelector('#fontFamily-cell').onChanged(value => {
				this.element.querySelectorAll('.user-recommended').forEach(element => {
					element.css({ fontFamily: value });
				});
			});

			layoutInfoRowPane.querySelector('#fontFamily-cell').onChanged(value => {
				this.element.querySelectorAll('.power-text').forEach(element => {
					element.css({ fontFamily: value });
				});
			});

			layoutHeaderRowPane.querySelector('#fontFamily-cell').onChanged(value => {
				this.element.querySelectorAll('.user-header').forEach(element => {
					element.css({ fontFamily: value });
				});
			});

			let showRecommended = layoutRecommendedRowPane.querySelector('#toggle-cell');
			showRecommended.addEventListener('change', e => {

				switch (showRecommended.value.toLowerCase()) {
					case "hide":
						this.element.querySelectorAll('.user-recommended').forEach(element => {
							element.style.display = "none";
						});
						break;
					case "show":
						this.element.querySelectorAll('.user-recommended').forEach(element => {
							element.style.display = "block";
						});
						break;
					default:
						this.element.querySelectorAll('.user-recommended').forEach(element => {
							element.style.display = "none";
						});
				}
			});

			let showInfo = layoutInfoRowPane.querySelector('#toggle-cell');
			showInfo.addEventListener('change', e => {

				switch (showInfo.value.toLowerCase()) {
					case "hide":
						this.element.querySelectorAll('.power-text').forEach(element => {
							element.style.display = "none";
						});
						break;
					case "show":
						this.element.querySelectorAll('.power-text').forEach(element => {
							element.style.display = "block";
						});
						break;
					default:
						this.element.querySelectorAll('.power-text').forEach(element => {
							element.style.display = "none";
						});
				}
			});

			let showHeader = layoutHeaderRowPane.querySelector('#toggle-cell');
			showHeader.addEventListener('change', e => {

				switch (showHeader.value.toLowerCase()) {
					case "hide":
						this.element.querySelectorAll('.user-header').forEach(element => {
							element.style.display = "none";
						});
						break;
					case "show":
						this.element.querySelectorAll('.user-header').forEach(element => {
							element.style.display = "block";
						});
						break;
					default:
						this.element.querySelectorAll('.user-header').forEach(element => {
							element.style.display = "none";
						});
				}
			});

			let userRowHandler = (userRowPane, userRowDom) => {
				let iconParent = userRowPane.querySelector('#image-cell').parentNode;
				this.uploadImage({ parent: iconParent }, (image) => {
					iconParent.querySelector('#image-cell').src = image.src;
					this.element.querySelector('.crater-power-image').src = image.src;
				});
				userRowPane.querySelector('#header-cell').onChanged(value => {
					userRowDom.querySelector('.user-header').innerHTML = value;
				});

				userRowPane.querySelector('#info-text-cell').onChanged(value => {
					userRowDom.querySelector('.power-text').innerHTML = value;
				});

				userRowPane.querySelector('#recommended-cell').onChanged(value => {
					userRowDom.querySelector('.user-recommended').innerHTML = value;
				});

				userRowPane.querySelector('#button-cell').onChanged(value => {
					userRowDom.querySelector('.user-button').innerHTML = value;
				});

			};

			let paneItems = this.paneContent.querySelectorAll('.crater-power-user-pane');
			paneItems.forEach((userRow, position) => {
				userRowHandler(userRow, userListRow[position]);
			});

		}

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
			let draftPower = this.sharePoint.properties.pane.content[this.key].settings.myPowerBi;

			if (this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.changed) {
				if (this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.groupId !== 'none') {
					this.getEmbedToken({ accessToken: draftPower.accessToken, groupID: draftPower.groupId, reportID: draftPower.reportId, generateUrl: draftPower.tokenEmbed }).then(response => {
						this.element.querySelector('.crater-power-container').css({ width: draftPower.width, height: draftPower.height });
						this.element.querySelector('#renderContainer').css({ width: '100%', height: draftPower.height });
						this.embedPower({ accessToken: response, type: draftPower.embedType, embedUrl: draftPower.embedUrl });
					});
				} else {
					if (draftPower.embedType === 'report') {
						if (this.element.querySelector('#renderContainer')) this.element.querySelector('#renderContainer').remove();
						let powerContainer = this.element.querySelector('.crater-power-container') as any;
						powerContainer.makeElement({
							element: 'div', attributes: { id: 'renderContainer' }, children: [
								{
									element: 'div', attributes: { id: 'render-error' }, children: [
										{ element: 'p', attributes: { id: 'render-text' } }
									]
								}
							]
						});
						let render = powerContainer.querySelector('#renderContainer');
						const iframeURL = `${draftPower.embedUrl}&autoAuth=true&ctid=${draftPower.tenantID}&filterPaneEnabled=${draftPower.showFilter}&navContentPaneEnabled=${draftPower.showNavContent}&pageName=${draftPower.namePage}`;
						render.makeElement({
							element: 'div', attributes: { class: 'power-iframe' }, children: [
								{
									element: 'iframe', attributes: { width: draftPower.width, height: draftPower.height, src: iframeURL, frameborder: "0", allowFullScreen: "true" }
								}
							]
						});
						this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.embedded = false;

						render.querySelector('#render-error').style.display = "none";
						render.style.display = "block";
					}
				}
			} else {
				if (this.sharePoint.properties.pane.content[this.key].settings.myPowerBi.groupId !== 'none') {
					this.element.querySelector('.crater-power-container').css({ width: draftPower.width, height: draftPower.height });
					this.element.querySelector('#renderContainer').css({ width: '100%', height: draftPower.height });
				} else {
					let powerIframe = this.element.querySelector('#renderContainer').querySelector('iframe');
					powerIframe.width = draftPower.width;
					powerIframe.height = draftPower.height;
				}

			}
		});
	}
}

{
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