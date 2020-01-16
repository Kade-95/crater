import { ElementModifier, func, ColorPicker, Connection } from '.';
import FroalaEditor from 'froala-editor';
import 'froala-editor/js/plugins/align.min.js';

require('./../styles/connection.css');
require('./../styles/form.css');
require('./../styles/events.css');
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
require('./../../../node_modules/froala-editor/css/froala_editor.pkgd.min.css');

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
		element.addEventListener('mouseenter', event => {
			if (element.hasAttribute('data-key') && this.sharePoint.inEditMode()) {
				element.querySelector('.webpart-options').css({ display: 'unset' });
			}
		});

		element.addEventListener('mouseleave', event => {
			if (element.hasAttribute('data-key')) {
				element.querySelector('.webpart-options').css({ display: 'none' });
			}
		});

		element.querySelectorAll('.keyed-element').forEach(keyedElement => {
			keyedElement.addEventListener('mouseenter', event => {
				if (keyedElement.hasAttribute('data-key') && this.sharePoint.inEditMode()) {
					keyedElement.querySelector('.webpart-options').css({ display: 'unset' });
				}
			});

			keyedElement.addEventListener('mouseleave', event => {
				if (keyedElement.hasAttribute('data-key')) {
					keyedElement.querySelector('.webpart-options').css({ display: 'none' });
				}
			});
		});
	}

	//generate webpart key
	private generateKey() {
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

		if (!func.isset(params.options)) params.options = ['Edit', 'Delete'];

		let options = this.webPartOptions({ options: params.options, title: params.attributes['data-type'] });
		delete params.options;

		let element = this.elementModifier.createElement(params);
		element.prepend(options);

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
						element: 'input', name: 'Color', attributes: {}, value: params.columns[i].querySelector('.crater-carousel-text').css().color || ''
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
					element: 'input', name: 'Color', attributes: {}
				}),
				this.elementModifier.cell({
					element: 'input', name: 'BackgroundColor', attributes: {}
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
		return events;
	}

	private rendered(params) {

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
								element: 'input', name: 'backgroundcolor', value: this.element.querySelector('.crater-events-title').css()['background-color']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'color', value: this.element.querySelector('.crater-events-title').css().color
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

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		//on save clicked save the webpart settings and re-render
		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());
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
						element: 'input', name: 'FontColor',
					}),
					this.elementModifier.cell({
						element: 'input', name: 'BackgroundColor',
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
								element: 'input', name: 'BackgroundColor'
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
			element: 'div', attributes: { class: 'crater-tab crater-component', 'data-type': 'tab' }, options: ['append', 'edit', 'delete'], children: [
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
					name: keyedElement.dataset.type,
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
				tabContentRowPane.dataset.title = value;
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
			{ image: 'https://ipigroup.sharepoint.com/sites/ShortPointTrial/SiteAssets/photo-1542178036-2e5efe4d8f83.jpg', text: 'text0' },
			{ image: 'https://ipigroup.sharepoint.com/sites/ShortPointTrial/SiteAssets/application-3426397_1920.jpg', text: 'text1' },
			{ image: 'https://ipigroup.sharepoint.com/sites/ShortPointTrial/SiteAssets/l1.jpg', text: 'text2' }
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
					this.elementModifier.createElement({ element: 'img', attributes: { src: image, alt: 'Not Found' } }),
					this.elementModifier.createElement({ element: 'p', attributes: { class: 'crater-slide-quote' }, text: params.source[j].text })
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
		this.startSlide();

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
				img.css({ height: this.element.position().height + 'px' });
			});
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
				element: 'div', attributes: { class: 'card', style: { margin: '1em', display: 'block' } }, sync: true, children: [
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
							})
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
				style: { border: '1px solid #bbbbbb', margin: '.5em 0em' }, class: 'list-row-pane row'
			},
			children: [
				this.paneOptions({ options: ['AB', 'AA', 'D'], owner: 'crater-slide-content-row' }),
				this.elementModifier.cell({
					element: 'img', name: 'Image', attributes: {}, dataAttributes: { class: 'crater-icon', src: this.sharePoint.images.append }
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Quote', value: 'quote'
				}),
			]
		});

		let slidePrototype = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'crater-slide', style: { display: 'none', opacity: 0 } }, children: [
				this.elementModifier.createElement({ element: 'img', attributes: { src: 'image', alt: 'Not Found' } }),
				this.elementModifier.createElement({ element: 'p', attributes: { class: 'crater-slide-quote' }, text: 'quote' })
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

			slides.append(newSlide);//c
			this.paneContent.querySelector('.list-pane').append(newListRow);

			listRowHandler(newListRow, newSlide);
		});

		this.paneContent.querySelectorAll('.crater-slide-row-pane').forEach((listRow, position) => {
			listRowHandler(listRow, slideListRows[position]);
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
								element: 'input', name: 'backgroundcolor', value: this.element.querySelector('.crater-list-title').css()['background-color']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'color', value: this.element.querySelector('.crater-list-title').css().color
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

			let update = this.paneContent.makeElement({
				element: 'div', attributes: { class: 'card update-pane update' }, children: [
					{
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Update'
							})
						]
					},
					{ element: 'label', text: 'The list to fetch must contain Title, Job, Link and Image ' },

					{
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'Site', attributes: {}, dataAttributes: {}
							}),
							this.elementModifier.cell({
								element: 'input', name: 'List', attributes: {}, dataAttributes: {}
							}),

							{
								element: 'div', attributes: { class: 'crater-center' }, children: [
									{ element: 'button', attributes: { class: 'btn', id: 'update-source', style: { display: 'flex', alignSelf: 'center' } }, text: 'Update Source' }
								]
							}
						]
					}
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

			this.paneContent.append(this.generatePaneContent({tiles}));

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
						element: 'input', name: 'Color', value: func.isset(params.tiles[i].css().color) ? params.tiles[i].css().color : this.color
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Background', value: func.isset(params.tiles[i].css()['background-color']) ? params.tiles[i].css()['background-color'] : this.backgroundColor
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
					element: 'input', name: 'Color', value: this.color
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Background', value: this.backgroundColor
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
					this.elementModifier.createElement({
						element: 'img', attributes: { class: 'crater-counter-content-column-image', src: count.image }
					}),
					this.elementModifier.createElement({
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
					})
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

				currentContent = this.element.makeElement({ element: 'div', attributes: { class: 'crater-counter-content', style: { 'gridTemplateColumns': `repeat(${columns}, 1fr)` } } });
			}

			currentContent.append(counter);
			counter.querySelector('.crater-counter-content-column-image').css({ height: this.backgroundHeight, width: this.backgroundWidth });

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

		this.height = func.isset(this.sharePoint.properties.pane.content[this.key].settings.height)
			? this.sharePoint.properties.pane.content[this.key].settings.height
			: this.height;

		if (this.height.toString().indexOf('px') == -1) this.height += 'px';

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
								element: 'input', name: 'BackgroundWidth', value: this.sharePoint.properties.pane.content[this.key].settings.backgroundWidth || ''
							}),
							this.elementModifier.cell({
								element: 'input', name: 'BackgroundHeight', value: this.sharePoint.properties.pane.content[this.key].settings.backgroundHeight || ''
							}),
							this.elementModifier.cell({
								element: 'select', name: 'BackgroundPosition', options: ['Left', 'Right']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Height', value: this.sharePoint.properties.pane.content[this.key].settings.height || ''
							})
						]
					})
				]
			});
		}

		// upload the settings

		this.paneContent.querySelector('#Duration-cell').value = this.sharePoint.properties.pane.content[this.key].settings.duration || '';

		this.paneContent.querySelector('#Columns-cell').value = this.sharePoint.properties.pane.content[this.key].settings.columns || '';

		this.paneContent.querySelector('#BackgroundPosition-cell').value = this.sharePoint.properties.pane.content[this.key].settings.backgroundPosition || '';

		this.paneContent.querySelector('#BackgroundWidth-cell').value = this.sharePoint.properties.pane.content[this.key].settings.backgroundWidth || '';

		this.paneContent.querySelector('#BackgroundHeight-cell').value = this.sharePoint.properties.pane.content[this.key].settings.backgroundHeight || '';

		this.paneContent.querySelector('#Height-cell').value = this.sharePoint.properties.pane.content[this.key].settings.height || '';

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
						element: 'input', name: 'Color', value: func.isset(params.counters[i].css().color) ? params.counters[i].css().color : this.color
					}),
					this.elementModifier.cell({
						element: 'input', name: 'Background', value: func.isset(params.counters[i].css()['background-color']) ? params.counters[i].css()['background-color'] : this.backgroundColor
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
					element: 'input', name: 'Color'
				}),
				this.elementModifier.cell({
					element: 'input', name: 'Background', value: this.backgroundColor
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

			this.sharePoint.properties.pane.content[this.key].settings.height = this.paneContent.querySelector('#Height-cell').value;

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

		this.element.addEventListener('click', event => {
			if (event.target.classList.contains('crater-ticker-news')) {
				event.preventDefault();
				let source = event.target.href;
				let openAt = this.sharePoint.properties.pane.content[this.key].settings.view;
				if (openAt.toLowerCase() == 'pop up') {
					this.element.append(this.elementModifier.popUp({ source, close: this.sharePoint.images.close }));
				}
				else if (openAt.toLowerCase() == 'new window') {
					window.open(source);
				}
				else {
					window.open(source, '_self');
				}
			}
		});
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
								element: 'input', name: 'BackgroundColor', value: this.element.querySelector('.crater-ticker-title').css()['background-color']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'TextColor', value: this.element.querySelector('.crater-ticker-title').css()['color']
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
			element: 'div', attributes: { class: 'crater-component', style: { display: 'block', minHeight: '100px', width: '100%' }, 'data-type': 'crater' }, options: ['Edit', 'Delete'], children: [
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
				element: 'section', attributes: { class: 'crater-section', 'data-type': 'section', style: {minHeight: params.height} }, options: ['Append', 'Edit', 'Delete'], type: 'crater-section', alignOptions: 'right', children: [
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
								element: 'input', name: 'color'
							}),
							this.elementModifier.cell({
								element: 'input', name: 'backgroundcolor'
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
								element: 'input', name: 'color', attributes: { type: 'number', min: 1 }, value: ''
							}),
							this.elementModifier.cell({
								element: 'input', name: 'backgroundcolor', attributes: { type: 'number', min: 1 }, value: ''
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
								element: 'input', name: 'bordercolor', attributes: { type: 'text' }, value: ''
							}),
							this.elementModifier.cell({
								element: 'input', name: 'borderstyle', attributes: { type: 'text' }, value: ''
							})
						]
					})
				]
			});

			this.paneContent.makeElement({
				element: 'div', attributes: { class: 'card update-pane' }, children: [
					{
						element: 'div', attributes: { class: 'card-title' }, children: [
							this.elementModifier.createElement({
								element: 'h2', attributes: { class: 'title' }, text: 'Update Table'
							})
						]
					},
					{ element: 'label', text: "The Data to get should be seperated with a comma[eg. Title, Name]" },
					{
						element: 'div', attributes: { class: 'row' }, children: [
							this.elementModifier.cell({
								element: 'input', name: 'Site', attributes: {}, dataAttributes: { value: this.sharePoint.properties.pane.content[this.key].settings.sourceSite || '' }
							}),
							this.elementModifier.cell({
								element: 'input', name: 'List', attributes: {}, dataAttributes: { value: this.sharePoint.properties.pane.content[this.key].settings.sourceList || '' }
							}),
							this.elementModifier.cell({
								element: 'input', name: 'Data', attributes: {}, dataAttributes: { value: this.sharePoint.properties.pane.content[this.key].settings.sourceData || '' }
							}),

							{
								element: 'div', attributes: { class: 'crater-center' }, children: [
									{ element: 'button', attributes: { class: 'btn', id: 'update-source', style: { display: 'flex', alignSelf: 'center' } }, text: 'Update Source' }
								]
							}
						]
					}
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
			element: 'div', attributes: { class: 'crater-panel crater-component', 'data-type': 'panel' }, options: ['Append', 'Edit', 'Delete'], children: [
				{ element: 'p', attributes: { class: 'crater-panel-title' }, text: 'Panel Title' },
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
								element: 'input', name: 'backgroundcolor', value: this.element.querySelector('.crater-panel-title').css()['background-color']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'color', value: this.element.querySelector('.crater-panel-title').css().color
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
								element: 'select', name: 'position', options: ['flex-start', 'flex-end', 'Center']
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

		titlePane.querySelector('#height-cell').onChanged(height => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-panel-title').css({ height });
		});

		titlePane.querySelector('#width-cell').onChanged(width => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-panel-title').css({ width });
		});

		titlePane.querySelector('#position-cell').onChanged(alignSelf => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-panel-title').css({ alignSelf });
		});

		titlePane.querySelector('#title-cell').onChanged(value => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-panel-title').innerText = value;
		});

		let backgroundColorCell = titlePane.querySelector('#backgroundcolor-cell').parentNode;
		this.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.querySelector('#backgroundcolor-cell') }, (backgroundColor) => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-panel-title').css({ backgroundColor });
			backgroundColorCell.querySelector('#backgroundcolor-cell').value = backgroundColor;
			backgroundColorCell.querySelector('#backgroundcolor-cell').setAttribute('value', backgroundColor);
		});

		let colorCell = titlePane.querySelector('#color-cell').parentNode;
		this.pickColor({ parent: colorCell, cell: colorCell.querySelector('#color-cell') }, (color) => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-panel-title').css({ color });
			colorCell.querySelector('#color-cell').value = color;
			colorCell.querySelector('#color-cell').setAttribute('value', color);
		});

		titlePane.querySelector('#fontsize-cell').onChanged(value => {
			this.sharePoint.properties.pane.content[this.key].draft.dom.querySelector('.crater-panel-title').css({ fontSize: value });
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
								element: 'input', name: 'Color',
							}),
							this.elementModifier.cell({
								element: 'input', name: 'BackgroundColor',
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
								element: 'input', name: 'Color',
							}),
							this.elementModifier.cell({
								element: 'input', name: 'BackgroundColor',
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
								element: 'input', name: 'backgroundcolor', value: this.element.querySelector('.crater-datelist-title').css()['background-color']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'color', value: this.element.querySelector('.crater-datelist-title').css().color
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
								element: 'input', name: 'dayColor', value: this.element.querySelector('.crater-datelist-content-item-date-day').css()['color']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'monthColor', value: this.element.querySelector('.crater-datelist-content-item-date-month').css()['color']
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
								element: 'input', name: 'titleColor', value: this.element.querySelector('.crater-datelist-content-item-text-main').css()['color']
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
								element: 'input', name: 'subtitleColor', value: this.element.querySelector('.crater-datelist-content-item-text-subtitle').css().color
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
								element: 'input', name: 'bodyColor', value: this.element.querySelector('.crater-datelist-content-item-text-body').css()['color']
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

	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.params = params;
	}

	public render(params) {

		let mapDiv = this.createKeyedElement({
			element: 'div', attributes: { class: 'crater-map', 'data-type': 'map' }, children: [
				{ element: 'div', attributes: { class: 'crater-map-div', id: 'crater-map-div' } }
			]
		});
		this.key = mapDiv.dataset.key;

		this.sharePoint.properties.pane.content[this.key].settings = { myMap: { lat: -34.067, lng: 150.067, zoom: 4, markerChecked: true, color: '' } };

		window['initMap'] = this.initMap;

		mapDiv.makeElement({
			element: 'script', attributes: { defer: '', async: '', src: 'https://maps.googleapis.com/maps/api/js?key=AIzaSyDdGAHe_9Ghatd4wZjyc3hRdirIQ1ttcv0&callback=initMap' }
		});

		return mapDiv;
	}

	public initMap = () => {
		let mapLocation = {
			lat: -34.067,
			lng: 150.067
		};

		//@ts-ignore
		let map = new google.maps.Map(document.querySelector('#crater-map-div'), {
			center: mapLocation,
			zoom: 4,
			mapTypeControlOptions: {
				mapTypeIds: ['roadmap', 'satellite', 'hybrid', 'terrain',
					'styled_map']
			}
		});

		//@ts-ignore
		var marker = new google.maps.Marker({
			position: mapLocation,
			map
		});

	}

	public rendered(params) {
	}

	public setUpPaneContent(params) {
		let key = params.element.dataset['key'];
		this.element = this.sharePoint.properties.pane.content[key].draft.dom;
		this.paneContent = this.elementModifier.createElement({
			elemen: 'div', attributes: { class: 'crater-property-content' }
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
								element: 'h2', attributes: { class: 'title' }, text: 'Personalise Map'
							})
						]
					}),
					this.elementModifier.createElement({
						element: 'div', attributes: { class: 'row' }, children: [
							{
								element: 'div', attributes: { class: 'message-note' }, children: [
									{ element: 'span', text: 'Note: Clear the color input field to reset the map color to default' }
								]
							}]
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
								element: 'input', name: 'color', value: ''
							}),
							this.elementModifier.cell({
								element: 'input', name: 'width', value: this.element.querySelector('#crater-map-div').css()['width']
							}),
							this.elementModifier.cell({
								element: 'input', name: 'height', value: this.element.querySelector('#crater-map-div').css()['height']
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


			let initMap = () => {
				let newMap = {
					lat: this.sharePoint.properties.pane.content[this.key].settings.myMap.lat,
					lng: this.sharePoint.properties.pane.content[this.key].settings.myMap.lng
				};

				const styles = (mapColor !== '') ? mapStyles : '';

				//@ts-ignore
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
			};

			window['initMap'] = initMap;

			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;
			this.element.querySelector('#crater-map-div').innerHTML = '';
			this.element.removeChild(this.element.querySelector('script'));
			this.element.makeElement({
				element: 'script', attributes: { defer: '', async: '', src: 'https://maps.googleapis.com/maps/api/js?key=AIzaSyDdGAHe_9Ghatd4wZjyc3hRdirIQ1ttcv0&callback=initMap' }
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
			elemen: 'div', attributes: { class: 'crater-property-content' }
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
			element: 'div', attributes: { class: 'crater-youtube', 'data-type': 'youtube' }, children: [
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
			elemen: 'div', attributes: { class: 'crater-property-content' }
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
	public newURL: any;
	public element: any;
	public paneContent: any;

	constructor(params) {
		super({ sharePoint: params.sharePoint });
		this.sharePoint = params.sharePoint;
		this.params = params;
	}

	public render(params) {
		if (!func.isset(params.url)) params.url = 'https://facebook.com/ipisolutions';

		let facebookDiv = this.createKeyedElement({
			element: 'div', attributes: { class: 'crater-facebook crater-component', 'data-type': 'facebook' }
		});

		this.sharePoint.properties.pane.content[facebookDiv.dataset.key].settings.url = params.url;

		return facebookDiv;
	}

	public rendered(params) {
		this.element = params.element;

		if (!func.isnull(this.element.querySelector('.crater-facebook-content'))) this.element.querySelector('.crater-facebook-content').remove();

		let url = this.sharePoint.properties.pane.content[this.element.dataset.key].settings.url;

		let facebookContent = this.elementModifier.createElement({
			element: 'div', attributes: { class: "crater-facebook-content" }, children: [
				{ element: 'div', attributes: { id: 'fb-root' } },
				{
					element: 'script', attributes: { src: 'https://connect.facebook.net/en_US/sdk.js#xfbml=1&version=v5.0&appId=541045216450969&autoLogAppEvents=1', async: 'false', defer: true, crossorigin: 'anonymous' }
				},
				{
					element: 'div', attributes: {
						class: 'fb-page',
						'data-href': url, 'data-tabs': 'timeline,messages,events', 'data-width': '', 'data-height': '', 'data-small-header': 'false', 'data-adapt-container-width': 'true', 'data-hide-cover': "false", 'data-show-facepile': "true"
					}
				}
			]
		});

		this.element.append(facebookContent);
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

		let pageUrl = this.paneContent.querySelector('#pageUrl-cell');
		pageUrl.onChanged();

		this.paneContent.addEventListener('mutated', event => {
			this.sharePoint.properties.pane.content[this.key].draft.pane.content = this.paneContent.innerHTML;
			this.sharePoint.properties.pane.content[this.key].draft.html = this.sharePoint.properties.pane.content[this.key].draft.dom.outerHTML;
		});

		this.paneContent.getParents('.crater-edit-window').querySelector('#crater-editor-save').addEventListener('click', event => {
			this.element.innerHTML = this.sharePoint.properties.pane.content[this.key].draft.dom.innerHTML;

			this.element.css(this.sharePoint.properties.pane.content[this.key].draft.dom.css());

			this.sharePoint.properties.pane.content[this.key].content = this.paneContent.innerHTML;

			this.sharePoint.properties.pane.content[this.key].settings.url = pageUrl.value;

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
									src: this.element.querySelector('.beforeAfter-contents').querySelector('.beforeImage').src
								}
							}),
							this.elementModifier.cell({
								element: 'img',
								name: 'after',
								dataAttributes: {
									style: { width: '400px', height: '400px' },
									src: this.element.querySelector('.beforeAfter-contents').querySelector('.crater-after').querySelector('.afterImage').src
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

{
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