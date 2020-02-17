import { Version, Environment, EnvironmentType, DisplayMode } from '@microsoft/sp-core-library';
import { ElementModifier, func, PropertyPane, CraterWebParts, Connection, ColorPicker, BaseClientSideWebPart, Images } from "../../externals/scr";

export interface ICraterProps {
	dom: any;
	pane: any;
	states: any;
}

export default class Crater extends BaseClientSideWebPart<ICraterProps> {
	public elementModifier = new ElementModifier({ sharePoint: this });
	public domContent: any;
	public app: any;
	public propertyPane: any = new PropertyPane({ sharePoint: this });
	public saved: boolean = false;
	public editted: boolean = false;
	public changingState: boolean = false;
	public savedWebPart: any;
	public dontSave: boolean = false;
	public craterWebparts = new CraterWebParts({ sharePoint: this });
	public displayPanelWindow: any;
	public displayPanelWindowExpanded: any = false;
	public pasteActive = false;
	public pasteElement: any;

	public images = Images;

	public connection: any = new Connection({ sharepoint: this });

	public render(): void {
		this.connection.context = this.context;

		if (!this.renderedOnce) {
			this.app = this.elementModifier.createElement({
				element: 'div', attributes: { class: 'crater', id: 'webpart-container', style: { width: '100%', zIndex: '1' } }
			}).monitor();

			if (this.properties.dom.generated && this.properties.dom.content != '') {// check if webpart has been created before
				this.domContent = this.elementModifier.createElement(this.properties.dom.content);
			} else {
				this.domContent = this.appendWebpart(this.app, 'crater');
				this.domContent.find('#undo-me').css({ opacity: 0.3 });
				this.domContent.find('#redo-me').css({ opacity: 0.3 });
				// Create single section for webpart
			}

			//set the base webpart properties      

			this.domElement.appendChild(this.app);
			this.app.appendChild(this.domContent);
			this.properties.dom.content = this.app.innerHTML;
			this.properties.dom.generated = true;

			//start running all webparts rendered
			this.runAll();

			if (this.properties.states.data.length == 0) {
				this.properties.states.data[0] = this.app.innerHTML;
			}

			this.app.findAll('.crater-display-panel').forEach(element => {
				element.remove();
			});

			this.app.addEventListener('mutated', event => {
				//check for changes
				if (this.dontSave) {
					this.dontSave = false;
				}
				else if (!this.changingState) {
					this.properties.dom.content = this.app.innerHTML;
					if (this.saved || this.editted) {
						this.saveCrater();
					}
				} else {
					this.changingState = false;
				}
			});

			if (!func.isnull(this.domContent)) {
				this.app.addEventListener('click', event => {
					let element = event.target;
					if (!(element.classList.contains('crater-display-panel') || element.getParents('.crater-display-panel') || element.classList.contains('new-component'))) {
						for (let displayPanel of this.app.findAll('.crater-display-panel')) {
							displayPanel.remove();
						}
					}

					if (this.inEditMode()) { //check if in edit mode
						if (element.id == 'edit-me') {//if edit is clicked
							this.propertyPane.render(element.getParents('data-key'));
						}
						else if (element.id == 'append-me') {//if append is clicked
							this.addWebpart(element);
						}
						else if (element.id == 'delete-me') {// if delete is clicked
							this.deleteWebpart(element);
						}
						else if (element.id == 'clone-me') {
							let choose = this.elementModifier.choose({ note: 'What do you want to do?', options: ["Copy", "Clone"] });

							this.app.append(choose.display);
							choose.choice.then((res: any) => {
								if (res.toLowerCase() == 'clone') {
									this.cloneWebpart(element);
								}
								else if (res.toLowerCase() == 'copy') {
									this.pasteActive = true;
									this.pasteElement = element.getParents('.crater-component');
								}
							});
						}
						else if (element.id == 'paste-me') {
							this.pasteWebpart(element);
						}
						else if (element.id == 'undo-me') {
							this.undoWebpart(element);
						}
						else if (element.id == 'redo-me') {
							this.redoWebpart(element);
						}
					}

					if (element.nodeName == 'A' && element.hasAttribute('href')) {
						event.preventDefault();
						this.openLink(element);
					}
				});
				this.initializeCrater();
				this.onWindowResized();
			}
		}
	}

	private runAll() {
		this.app.findAll('.keyed-element').forEach(element => {
			if (element.hasAttribute('data-type')) {
				let type = element.dataset.type;
				this.craterWebparts[type]({ action: 'rendered', element, sharePoint: this });
			}
		});
	}

	private openLink(element) {
		let source = element.href;
		let webpart = element;
		if (!(element.classList.contains('crater-component'))) webpart = element.getParents('.crater-component');
		let openAt = this.properties.pane.content[webpart.dataset.key].settings.view || 'same window';
		if (openAt.toLowerCase() == 'pop up') {
			let popUp = this.elementModifier.popUp({ source, close: this.images.close, maximize: this.images.maximize, minimize: this.images.minimize });
			webpart.append(popUp);
		}
		else if (openAt.toLowerCase() == 'new window') {
			window.open(source);
		}
		else {
			window.open(source, '_self');
		}
	}

	private onWindowResized() {
		//remove all editwindows 
		window.onresize = () => {
			//reset the size of the editwindow to match the size of the screen
			let editWindow = this.app.find('.crater-edit-window');
			if (!func.isnull(editWindow)) {
				editWindow.position({ height: window.innerHeight, width: window.innerWidth });
				editWindow.find('.crater-editor').css({
					height: `${8 * window.innerHeight / 10}px`,
					width: `${9 * window.innerWidth / 10}px`,
					marginTop: `${0.5 * window.innerHeight / 10}px`,
					marginLeft: `${0.5 * window.innerWidth / 10}px`
				});
			}
		};
	}

	private initializeCrater() {
		this.app.findAll('.crater-edit-window').forEach(element => {
			element.remove();
		});
		this.app.findAll('.crater-pop-up').forEach(element => {
			element.remove();
		});
	}

	public appendWebpart(parent, webpart) {
		//fetch webpart and append it to the section || 
		let element = this.craterWebparts[webpart]({ action: 'render', sharePoint: this });
		parent.append(element);
		this.craterWebparts[webpart]({ action: 'rendered', element, sharePoint: this });

		if (func.isset(this.domContent)) {
			this.craterWebparts['crater']({ action: 'rendered', element: this.domContent, sharePoint: this });
		}
		return element;
	}

	protected get dataVersion(): Version {
		return Version.parse('1.0');
	}

	public displayPanel(selected) {

		let webparts = ['Panel', 'List', 'Slider', 'Counter', 'Tiles', 'News', 'Table', 'TextArea', 'Icons', 'Button', 'Count Down', 'Tab', 'Events', 'Carousel', 'Map', 'Date List', 'Instagram', 'Facebook', 'Youtube', 'Twitter', 'Before After', 'Event', 'Power', 'Employee Directory', 'Birthday'];

		this.displayPanelWindow = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'crater-display-panel' }
		});

		let controls = this.displayPanelWindow.makeElement({
			element: 'div', attributes: { class: 'display-pane-controls' }
		});

		//search box 
		this.displayPanelWindow.makeElement({
			element: 'input', attributes: { id: 'search-webpart', placeHolder: 'Search' }
		})
			.onChanged(value => {
				let foundWebParts = [];
				for (let i of webparts) {
					if (i.toLowerCase().indexOf(value.toLowerCase()) != -1) {
						foundWebParts.push(i);
					}
				}

				this.updateDisplayPaneWebPart({ webparts: foundWebParts });
			});

		this.displayPanelWindow.makeElement({
			element: 'div', attributes: { class: 'select-webparts', id: 'select-webpart' }
		});

		this.updateDisplayPaneWebPart({ webparts });

		controls.makeElement({
			element: 'img', attributes: { id: 'toggle', src: this.images.maximize, class: 'crater-icon display-pane-controls-button' }
		}).addEventListener('click', event => {
			event.target.classList.toggle('wide');
			if (event.target.classList.contains('wide')) {
				event.target.src = this.images.minimize;
			} else {
				event.target.src = this.images.maximize;
			}

			this.displayPanelWindow.classList.toggle('wide');
		});

		controls.makeElement({
			element: 'img', attributes: { id: 'close', src: this.images.close, class: 'crater-icon display-pane-controls-button' }
		}).addEventListener('click', event => {
			this.displayPanelWindow.remove();
		});


		this.displayPanelWindow.addEventListener('click', event => {
			let element = event.target;
			if (element.classList.contains('single-webpart') || !func.isnull(element.getParents('.single-webpart'))) {
				//   //select webpart to append
				if (!element.classList.contains('single-webpart')) element = element.getParents('.single-webpart');
				selected(element);
				this.displayPanelWindow.remove();
			}
		});

		return this.displayPanelWindow;
	}

	private updateDisplayPaneWebPart(params) {
		this.displayPanelWindow.find('#select-webpart').innerHTML = '';//clear window
		for (let single of params.webparts) {
			this.displayPanelWindow.find('#select-webpart').makeElement({
				element: 'div', attributes: { class: 'single-webpart', 'data-webpart': func.stringReplace(single.toLowerCase(), ' ', '') }, children: [
					this.elementModifier.createElement({//set the icon
						element: 'img', attributes: { class: 'image', src: this.images.append }
					}),
					this.elementModifier.createElement({
						element: 'a', attributes: { class: 'title' }, text: single//set the text
					})
				]
			});
		}
	}

	public inEditMode() {
		let editing = this.displayMode == DisplayMode.Edit;
		if (!editing) {
			this.app.findAll('.webpart-option').forEach(option => {
				option.show();
			});
		}
		return editing;
	}

	public get isLocal() {
		let local = Environment.type == EnvironmentType.Local;
		return local;
	}

	public addWebpart(element) {
		this.app.findAll('.crater-display-panel').forEach(panel => {
			panel.remove();
		});

		element.getParents('data-key').append(
			this.displayPanel(webpart => {
				let container = element.getParents('.crater-panel') || element.getParents('.crater-tab') || element.getParents('.crater-section');

				if (container.classList.contains('crater-section')) {
					this.appendWebpart(container.find('.crater-section-content'), webpart.dataset.webpart);
				} else if (container.classList.contains('crater-panel')) {
					this.appendWebpart(container.find('.crater-panel-content'), webpart.dataset.webpart);
				} else if (container.classList.contains('crater-tab')) {
					this.appendWebpart(container.find('.crater-tab-content'), webpart.dataset.webpart);
					this.craterWebparts['tab']({ action: 'rendered', element: container, sharePoint: this });
				}

				this.editted = true;
			})
		);
	}

	public deleteWebpart(element) {
		if (confirm("Do you want to continue with this action")) {//confirm deletion
			let key = element.getParents('data-key').dataset.key;
			if (element.getParents('data-key').outerHTML == this.domContent.outerHTML) {
				//if element is the base webpart
				this.domContent.getParents('.ControlZone').remove();
				this.properties.dom.content = 'Webpart Deleted';
			}
			else if (element.getParents('data-key').classList.contains('crater-section')) {
				//if element is a section
				element.getParents('data-key').remove();
				this.properties.pane.content[this.domContent.dataset['key']].settings.columns -= 1;
				let columns = this.properties.pane.content[this.domContent.dataset['key']].settings.columns;

				this.properties.pane.content[this.domContent.dataset['key']].settings.columnsSizes = `repeat(${columns} 1fr)`;
				this.craterWebparts['crater']({ action: 'rendered', element: this.domContent, sharePoint: this, resetWidth: true });
			}
			else {
				element = element.getParents('data-key');
				let tab = element.getParents('.crater-tab');
				element.remove();

				if (!func.isnull(tab)) {
					this.craterWebparts['tab']({ action: 'rendered', element: tab, sharePoint: this });
				}
				this.properties.dom.content = this.domContent.outerHTML;
			}
			this.editted = true;
		}
	}

	public cloneWebpart(element) {
		let webpart = element.getParents('.crater-component');
		let clone = webpart.cloneNode(true);

		let container = webpart.getParents('.crater-component');

		let newKey = this.craterWebparts.generateKey();
		clone.dataset.key = newKey;
		this.properties.pane.content[newKey] = this.properties.pane.content[webpart.dataset.key];

		if (container.classList.contains('crater-crater')) {
			this.properties.pane.content[container.dataset.key].settings.columns =
				this.properties.pane.content[container.dataset.key].settings.columns + 1;

			container.find('.crater-sections-container').css({ gridTemplateColumns: `repeat(${this.properties.pane.content[container.dataset.key].settings.columns}, 1fr)` });
		}

		webpart.after(clone);

		this.craterWebparts[clone.dataset.type]({ action: 'rendered', element: clone, sharePoint: this });

		this.craterWebparts[container.dataset.type]({ action: 'rendered', element: container, sharePoint: this });
		this.editted = true;
	}

	public pasteWebpart(element) {
		let clone = this.pasteElement.cloneNode(true);
		let container = element.getParents('.crater-container');

		let newKey = this.craterWebparts.generateKey();
		clone.dataset.key = newKey;
		this.properties.pane.content[newKey] = this.properties.pane.content[this.pasteElement.dataset.key];

		if (container.classList.contains('crater-section')) {
			container.find('.crater-section-content').append(clone);
		}
		else if (container.classList.contains('crater-panel')) {
			container.find('.crater-panel-content').append(clone);
		}
		else if (container.classList.contains('crater-tab')) {
			container.find('.crater-tab-content').append(clone);
		}

		this.craterWebparts[clone.dataset.type]({ action: 'rendered', element: clone, sharePoint: this });

		this.craterWebparts[container.dataset.type]({ action: 'rendered', element: container, sharePoint: this });

		this.pasteActive = false;
		this.editted = true;
	}

	public changeState() {
		this.changingState = true;
		this.app.innerHTML = this.properties.states.data[this.properties.states.currentPosition] || this.app.innerHTML;
		this.runAll();

		this.setCorrection();
	}

	public redoWebpart(element: any) {
		if (this.properties.states.data.length > this.properties.states.currentPosition + 1) {
			this.properties.states.currentPosition = this.properties.states.currentPosition + 1;
			this.changeState();
		}
	}

	public undoWebpart(element: any) {
		if (this.properties.states.currentPosition != 0) {
			this.properties.states.currentPosition = this.properties.states.currentPosition - 1;
			this.changeState();
		}
	}

	private saveCrater() {
		//show options of the keyed elements
		if (this.saved) {
			let type = this.savedWebPart.dataset.type;
			//start the re-running the webpart
			this.craterWebparts[type]({ action: 'rendered', element: this.savedWebPart, sharePoint: this });
			this.propertyPane.clearDraft(this.properties.pane.content[this.savedWebPart.dataset.key].draft);
		}
		this.properties.states.currentPosition = this.properties.states.currentPosition / 1 + 1;
		this.app.find('#undo-me').css({ opacity: 1 });
		for (let i in this.properties.states.data) {
			if (i < this.properties.states.currentPosition) continue;
			this.properties.states.data.pop(i);
		}
		this.properties.states.data.push(this.properties.dom.content);

		this.saved = false;
		this.editted = false;
	}

	private setCorrection(){
		if (this.properties.states.currentPosition == 0) {
			this.app.find('#undo-me').css({ opacity: 0.3 });
		} else {
			this.app.find('#undo-me').css({ opacity: 1 });
		}

		if (this.properties.states.currentPosition + 1 == this.properties.states.data.length) {
			this.app.find('#redo-me').css({ opacity: 0.3 });
		}
		else {
			this.app.find('#redo-me').css({ opacity: 1 });
		}
	}
}