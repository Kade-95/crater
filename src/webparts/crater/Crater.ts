import { Version, Environment, EnvironmentType, DisplayMode } from '@microsoft/sp-core-library';
import { ElementModifier, func, PropertyPane, CraterWebParts, Connection, ColorPicker, BaseClientSideWebPart, Images } from "../../externals/scr";

export interface ICraterProps {
	dom: any;
	pane: any;
}

export default class Crater extends BaseClientSideWebPart<ICraterProps> {
	public elementModifier = new ElementModifier({ sharePoint: this });
	public domContent: any;
	public app: any;
	public propertyPane: any = new PropertyPane({ sharePoint: this });
	public saved: boolean = false;
	public savedWebPart: any;
	public craterWebparts = new CraterWebParts({ sharePoint: this });
	public displayPanelWindow: any;
	public displayPanelWindowExpanded: any = false;

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
				// Create single section for webpart
			}

			//set the base webpart properties      

			this.domElement.appendChild(this.app);
			this.app.appendChild(this.domContent);
			this.properties.dom.content = this.app.innerHTML;
			this.properties.dom.generated = true;

			//start running all webparts rendered
			this.app.querySelectorAll('.keyed-element').forEach(element => {
				if (element.hasAttribute('data-type')) {
					let type = element.dataset.type;
					this.craterWebparts[type]({ action: 'rendered', element, sharePoint: this });
				}
			});

			this.app.querySelectorAll('.crater-display-panel').forEach(element => {
				element.remove();
			});

			this.app.addEventListener('mutated', event => {
				//check for changes
				this.properties.dom.content = this.app.innerHTML;
				if (this.saved) {
					//if saved 
					this.saved = false;
					//show options of the keyed elements
					let type = this.savedWebPart.dataset.type;
					//start the re-running the webpart
					this.craterWebparts[type]({ action: 'rendered', element: this.savedWebPart, sharePoint: this });
					this.propertyPane.clearDraft(this.properties.pane.content[this.savedWebPart.dataset.key].draft);
				}
			});

			if (!func.isnull(this.domContent)) {
				this.app.addEventListener('click', event => {
					let element = event.target;
					if (!(element.classList.contains('crater-display-panel') || element.getParents('.crater-display-panel') || element.classList.contains('new-component'))) {
						for (let displayPanel of this.app.querySelectorAll('.crater-display-panel')) {
							displayPanel.remove();
						}
					}

					if (this.inEditMode()) { //check if in edit mode
						if (element.id == 'edit-me') {//if edit is clicked
							this.propertyPane.render(element.getParents('data-key'));
						}
						else if (element.id == 'append-me') {//if append is clicked
							this.app.querySelectorAll('.crater-display-panel').forEach(panel => {
								panel.remove();
							});

							element.getParents('data-key').append(
								this.displayPanel(webpart => {
									let container = element.getParents('.crater-panel') || element.getParents('.crater-tab') || element.getParents('.crater-section');

									if (container.classList.contains('crater-section')) {
										this.appendWebpart(container.querySelector('.crater-section-content'), webpart.dataset.webpart);
									} else if (container.classList.contains('crater-panel')) {
										this.appendWebpart(container.querySelector('.crater-panel-content'), webpart.dataset.webpart);
									} else if (container.classList.contains('crater-tab')) {
										this.appendWebpart(container.querySelector('.crater-tab-content'), webpart.dataset.webpart);
										this.craterWebparts['tab']({ action: 'rendered', element: container, sharePoint: this });
									}
								})
							);
						}
						else if (element.id == 'delete-me') {// if delete is clicked
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
								delete this.properties.pane.content[key];
							}
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

	private openLink(element) {
		let source = element.href;
		let webpart = element;
		if (!(element.classList.contains('crater-component'))) webpart = element.getParents('.crater-component');
		let openAt = this.properties.pane.content[webpart.dataset.key].settings.view || 'same window';
		if (openAt.toLowerCase() == 'pop up') {
			let popUp = this.elementModifier.popUp({ source, close: this.images.close });
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
			let editWindow = this.app.querySelector('.crater-edit-window');
			if (!func.isnull(editWindow)) {
				editWindow.position({ height: window.innerHeight, width: window.innerWidth });
				editWindow.querySelector('.crater-editor').css({
					height: `${8 * window.innerHeight / 10}px`,
					width: `${9 * window.innerWidth / 10}px`,
					marginTop: `${0.5 * window.innerHeight / 10}px`,
					marginLeft: `${0.5 * window.innerWidth / 10}px`
				});
			}
		};
	}

	private initializeCrater() {
		this.app.querySelectorAll('.crater-edit-window').forEach(element => {
			element.remove();
		});
		this.app.querySelectorAll('.crater-pop-up').forEach(element => {
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

		let webparts = ['Panel', 'List', 'Slider', 'Counter', 'Tiles', 'News', 'Table', 'TextArea', 'Icons', 'Button', 'Count Down', 'Tab', 'Events', 'Carousel', 'Map', 'Date List', 'Instagram', 'Facebook', 'Before After', 'Youtube', 'Event', 'Power', 'Employee Directory'];

		this.displayPanelWindow = this.elementModifier.createElement({
			element: 'div', attributes: { class: 'crater-display-panel' }, text: 'Display'
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

		let controls = this.displayPanelWindow.makeElement({
			element: 'div', attributes: { class: 'display-pane-controls' }
		});

		controls.makeElement({
			element: 'button', attributes: { id: 'toggle' }, text: 'Toggle'
		}).addEventListener('click', event => {

			if (!this.displayPanelWindowExpanded) {
				this.displayPanelWindow.css({
					top: (innerHeight * 0.1) + 'px',
					right: (innerWidth * 0.1) + 'px',
				});

				this.displayPanelWindow.querySelector('#search-webpart').css({
					width: (innerWidth * 0.8) + 'px',
				});

				this.displayPanelWindow.querySelector('#select-webpart').css({
					width: (innerWidth * 0.8) + 'px',
					height: (innerHeight * 0.8 - this.displayPanelWindow.querySelector('.display-pane-controls').position().height) + 'px',
				});

				this.displayPanelWindow.querySelector('.display-pane-controls').css({
					width: (innerWidth * 0.8) + 'px',
				});

				this.displayPanelWindowExpanded = true;
			}
			else {
				this.displayPanelWindow.css({
					top: '0px',
					right: '0px',
				});

				this.displayPanelWindow.querySelector('#search-webpart').css({
					width: '290px',
				});

				this.displayPanelWindow.querySelector('#select-webpart').css({
					width: '290px',
					height: '300px',
				});

				this.displayPanelWindow.querySelector('.display-pane-controls').css({
					width: '300px',
				});

				this.displayPanelWindowExpanded = false;
			}
		});

		controls.makeElement({
			element: 'button', attributes: { id: 'close' }, text: 'Close'
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
		this.displayPanelWindow.querySelector('#select-webpart').innerHTML = '';//clear window
		for (let single of params.webparts) {
			this.displayPanelWindow.querySelector('#select-webpart').makeElement({
				element: 'div', attributes: { class: 'single-webpart', 'data-webpart': func.stringReplace(single.toLowerCase(), ' ', '') }, children: [
					this.elementModifier.createElement({//set the icon
						element: 'img', attributes: { class: 'image' }
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
			this.app.querySelectorAll('.webpart-option').forEach(option => {
				option.show();
			});
		}
		return editing;
	}
}