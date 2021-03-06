import { ElementModifier, Func, CraterWebParts, ColorPicker } from "./index";

class PropertyPane {
    public sharePoint: any;
    public paneContent: any;
    public paneStyle: any;
    public paneConnection: any;
    public elementModifier: any;
    public element: any;
    public editor: any;
    public craterWebparts: any;
    public currentWindow: any;
    public func = new Func();

    constructor(params) {
        this.sharePoint = params.sharePoint;
        this.elementModifier = new ElementModifier(params.sharePoint);
        this.craterWebparts = new CraterWebParts(params.sharePoint);
        this.paneContent = this.elementModifier.createElement({
            element: 'div',
            attributes: { class: 'crater-property-content' }
        }).monitor();
        this.paneStyle = this.elementModifier.createElement({
            element: 'div',
            attributes: { class: 'crater-property-style' }
        }).monitor();
        this.paneConnection = this.elementModifier.createElement({
            element: 'div',
            attributes: { class: 'crater-property-connection' }
        }).monitor();
    }

    public render(element): any {
        this.sharePoint.app.findAll('.webpart-options').forEach(option => {
            option.hide();
        });
        this.element = element;

        let key = this.element.dataset['key'];
        if (!this.func.isset(key)) {
            alert("This element cannot be editted");
            return;
        }

        if (!this.func.isset(this.sharePoint.storage.pane.content[key])) {
            this.sharePoint.storage.pane.content[key] = { content: '', styles: '', settings: {}, draft: { dom: '', html: '', pane: { content: '', styles: '' } } };
        }

        let content = [
            { name: 'Content', owner: 'Content' },
            { name: 'Styles', owner: 'Styles' },
        ];

        if (this.func.isset(this.element.dataset.connectible)) {
            content.push({ name: 'Connection', owner: 'Connection' });
        }

        let menus = this.elementModifier.menu({
            content,
            padding: '1em 0em'
        });

        this.editor = this.elementModifier.createElement({
            element: 'div', attributes: { class: 'crater-editor' }
        });

        let editWindow = this.elementModifier.createElement({
            element: 'div', attributes: {
                class: 'crater-edit-window'
            },
            children: [
                menus,
                this.editor
            ]
        });

        menus.addEventListener('click', event => {
            let item = event.target;
            if (!item.classList.contains('crater-menu-item')) item = item.getParents('.crater-menu-item');
            if (this.func.isnull(item)) return;

            if (item.dataset.owner == 'Content' && this.currentWindow != 'Content') {
                this.setUpContent(key);
            }
            else if (item.dataset.owner == 'Styles' && this.currentWindow != 'Styles') {
                this.setUpStyle(key);
            }
            else if (item.dataset.owner == 'Connection' && this.currentWindow != 'Connection') {
                this.setUpConnection(key);
            }

            this.currentWindow = item.dataset.owner;

            menus.findAll('.crater-menu-item').forEach(mItem => {
                mItem.cssRemove(['background-color']);
            });

            item.css({ backgroundColor: `var(--lighter-primary-color)` });
        });

        editWindow.makeElement({
            element: 'div', attributes: { style: { position: 'absolute', bottom: '0px', marginBottom: '1.1em', right: '5%' } }, children: [
                this.elementModifier.createElement({
                    element: 'button', attributes: { id: 'crater-editor-save', class: 'btn' }, text: 'Save'
                }),
                this.elementModifier.createElement({
                    element: 'button', attributes: { id: 'crater-editor-cancel', class: 'btn' }, text: 'Cancel'
                })
            ]
        })
            .addEventListener('click', event => {
                if (event.target.id == 'crater-editor-save') {//save is clicked
                    this.sharePoint.storage.pane.content[key].styles = this.paneStyle.innerHTML;

                    if (this.func.isset(this.sharePoint.storage.pane.content[key].draft.dom.dataset.backgroundImage)) {
                        this.element.setBackgroundImage(this.sharePoint.storage.pane.content[key].draft.dom.dataset.backgroundImage);
                    }

                    this.sharePoint.saved = true;
                    this.sharePoint.savedWebPart = this.element;

                    console.log('Crater edition saved');
                } else if (event.target.id == 'crater-editor-cancel') {//keep draft and exit
                    this.clearDraft(this.sharePoint.storage.pane.content[key].draft);
                    console.log('Crater edition cancelled');
                }
                editWindow.remove();
            });
        this.sharePoint.app.append(editWindow);

        this.setUpContent(key);

        return editWindow;
    }

    private setUpContent(key): any {
        // get webpart
        if (this.sharePoint.storage.pane.content[key].draft.html == '') {
            this.sharePoint.storage.pane.content[key].draft.dom = this.element.cloneNode(true);
            this.sharePoint.storage.pane.content[key].draft.html = this.sharePoint.storage.pane.content[key].draft.dom.outerHTML;
        }
        else {
            this.sharePoint.storage.pane.content[key].draft.dom = this.elementModifier.createElement(this.sharePoint.storage.pane.content[key].draft.html);
        }

        let draftDom = this.sharePoint.storage.pane.content[key].draft.dom;
        let type = draftDom.dataset.type;

        this.editor.innerHTML = '';
        this.editor.append(this.craterWebparts[type]({ action: 'setUpPaneContent', element: this.element, draft: draftDom, sharePoint: this.sharePoint }));

        this.craterWebparts[type]({ action: 'listenPaneContent', element: this.element, draft: draftDom, sharePoint: this.sharePoint });

        this.editor.findAll('.crater-color-picker').forEach(picker => {
            picker.remove();
        });
    }

    private setUpConnection(key): any {
        // get webpart
        let type = this.element.dataset.type;
        let metaData = [];

        let update = this.craterWebparts[type]({ action: 'update', element: this.element, sharePoint: this.sharePoint, options: [] });

        this.paneConnection = this.elementModifier.createElement({
            element: 'div',
            attributes: { class: 'crater-property-connection' },
        }).monitor();

        let getDisplayOptions = (connectionType) => {
            if (connectionType == 'Same Site') {
                return this.elementModifier.createForm({
                    title: 'Same Site Connection', attributes: { id: 'connection-form', class: 'form' },
                    contents: {
                        list: { element: 'input', attributes: { id: 'connection-list', name: 'List' } },
                    },
                    buttons: {
                        submit: { element: 'button', attributes: { id: 'create-connection', class: 'btn' }, text: 'Submit' },
                    }
                });
            }
            else if (connectionType == 'Other Sharepoint Site') {
                return this.elementModifier.createForm({
                    title: 'Another SharePoint Site Connection', attributes: { id: 'connection-form', class: 'form' },
                    contents: {
                        link: { element: 'input', attributes: { id: 'connection-link', name: 'Link' } },
                        list: { element: 'input', attributes: { id: 'connection-list', name: 'List' } },
                    },
                    buttons: {
                        submit: { element: 'button', attributes: { id: 'create-connection', class: 'btn' }, text: 'Submit' },
                    }
                });
            }
            else if (connectionType == 'RSS Feed') {
                return this.elementModifier.createForm({
                    title: 'RSS Feed Connection', attributes: { id: 'connection-form', class: 'form' },
                    contents: {
                        link: { element: 'input', attributes: { id: 'connection-link', name: 'Link' } },
                        count: { element: 'input', attributes: { id: 'connection-count', name: 'Count', type: 'number', min: 1 } }
                    },
                    buttons: {
                        submit: { element: 'button', attributes: { id: 'create-connection', class: 'btn' }, text: 'Submit' },
                    }
                });
            }
        };

        if (this.sharePoint.storage.pane.content[key].draft.pane.connection != '') {
            this.paneConnection.innerHTML = this.sharePoint.storage.pane.content[key].draft.pane.connection;
        }
        else if (this.sharePoint.storage.pane.content[key].styles != '') {
            this.paneConnection.innerHTML = this.sharePoint.storage.pane.content[key].connection;
        }
        else {
            this.paneConnection.makeElement({
                element: 'div', attributes: { class: 'crater-connection-content' }, children: [
                    {
                        element: 'section', attributes: { class: 'crater-connection-try' }, children: [
                            this.elementModifier.cell({ element: 'select', name: 'Type', options: ['Same Site', 'Other Sharepoint Site', 'RSS Feed'] }),
                            { element: 'div', attributes: { id: 'crater-connection-type-option' } }
                        ]
                    },
                    {
                        element: 'section', attributes: { class: 'crater-connection-get' }, children: [
                            update
                        ]
                    }
                ]
            });
        }

        let getWindow = this.paneConnection.find('.crater-connection-get');

        this.paneConnection.find('#Type-cell').onChanged(value => {
            this.paneConnection.find('#crater-connection-type-option').innerHTML = '';
            this.paneConnection.find('#crater-connection-type-option').makeElement({ element: getDisplayOptions(value) });
        });

        this.paneConnection.find('.form-error').innerHTML = '';

        this.paneConnection.addEventListener('click', event => {
            let target = event.target;

            let fetchData = (link, list, form, formError) => {
                if (!this.elementModifier.validateForm(form)) {
                    formError.textContent = 'Form not filled correctly';
                    return;
                }
                this.sharePoint.connection.find({ link, list, data: '' }).then(source => {
                    metaData = this.func.getObjectArrayKeys(source);
                    update = this.craterWebparts[type]({ action: 'update', element: this.element, sharePoint: this.sharePoint, options: metaData, source });
                    getWindow.innerHTML = '';
                    getWindow.append(update);

                    formError.textContent = 'Connected';
                });
            };

            if (target.id == 'create-connection') {
                event.preventDefault();
                let form = target.getParents('.form');
                let formError = form.find('.form-error');
                formError.css({ display: 'unset' });
                formError.textContent = 'Connecting...';

                let connectionType = this.paneConnection.find('#Type-cell').value;
                if (connectionType == 'Same Site') {
                    let link = this.sharePoint.connection.getSite();
                    let list = this.paneConnection.find('#connection-list').value;
                    fetchData(link, list, form, formError);
                }
                else if (connectionType == 'Other Sharepoint Site') {
                    let link = this.paneConnection.find('#connection-link').value;
                    let list = this.paneConnection.find('#connection-list').value;
                    fetchData(link, list, form, formError);
                }
                else if (connectionType == 'RSS Feed') {

                    if (!this.elementModifier.validateForm(form)) {
                        formError.textContent = 'Form not filled correctly';
                        return;
                    }

                    let url = `https://cors-anywhere.herokuapp.com/` + this.paneConnection.find('#connection-link').value;
                    let count = this.paneConnection.find('#connection-count').value;

                    this.sharePoint.connection.ajax({ method: 'GET', url }).then(result => {
                        formError.textContent = 'Connected';
                        let domParser = new DOMParser();
                        let doc = domParser.parseFromString(result, 'text/xml');
                        let items: any = doc.querySelectorAll('item');

                        if (items.length == 0) {
                            items = doc.querySelectorAll('entry');
                        }

                        let source = [];
                        for (let i = 0; i < items.length; i++) {
                            if (i == count) break;
                            let item = items[i];

                            let row = {};
                            let content = item.childNodes;
                            for (let j = 0; j < content.length; j++) {
                                row[content[j].nodeName] = content[j].textContent;
                            }
                            source.push(row);
                        }
                        metaData = this.func.getObjectArrayKeys(source);
                        update = this.craterWebparts[type]({ action: 'update', element: this.element, sharePoint: this.sharePoint, options: metaData, source });
                        getWindow.innerHTML = '';
                        getWindow.append(update);
                    });
                }
            }
        });

        this.editor.innerHTML = '';
        this.editor.append(this.paneConnection);

        this.paneConnection.addEventListener('mutated', event => {
            this.sharePoint.storage.pane.content[key].draft.pane.connection = this.paneConnection.innerHTML;
            this.sharePoint.storage.pane.content[key].draft.html = this.sharePoint.storage.pane.content[key].draft.dom.outerHTML;
        });
    }

    private setUpStyle(key): any {
        this.paneStyle = this.elementModifier.createElement({
            element: 'div',
            attributes: { class: 'crater-property-style' }
        }).monitor();

        if (this.sharePoint.storage.pane.content[key].draft.pane.styles != '') {
            this.paneStyle.innerHTML = this.sharePoint.storage.pane.content[key].draft.pane.styles;
        }
        else if (this.sharePoint.storage.pane.content[key].styles != '') {
            this.paneStyle.innerHTML = this.sharePoint.storage.pane.content[key].styles;
        }
        else {
            let paddings = { paddingTop: 'Padding Top', paddingLeft: 'Padding Left', paddingBottom: 'Padding Bottom', paddingRight: 'Padding Bottom' };
            let paddingBlock = this.sharePoint.craterWebparts.createStyleBlock({ children: paddings, title: "Paddings", element: this.element, options: { sync: true } });
            this.paneStyle.append(paddingBlock);

            let margins = { marginTop: 'Margin Top', marginLeft: 'Margin Left', marginBottom: 'Margin Bottom', marginRight: 'Margin Right' };
            let marginBlock = this.sharePoint.craterWebparts.createStyleBlock({ children: margins, title: "Margins", element: this.element, options: { sync: true } });
            this.paneStyle.append(marginBlock);

            let borders = { borderTop: 'Border Top', borderLeft: 'Border Left', borderBottom: 'Border Bottom', borderRight: 'Border Right' };
            let borderBlock = this.sharePoint.craterWebparts.createStyleBlock({ children: borders, title: "Borders", element: this.element, options: { sync: true } });
            this.paneStyle.append(borderBlock);

            let borderRadius = { borderTopLeftRadius: 'Top-Left Radius', borderBottomLeftRadius: 'Bottom-Left Radius', borderTopRightRadius: 'Top-Right Radius', borderBottomRightRadius: 'Bottom-Right Radius', };
            let borderRadiusBlock = this.sharePoint.craterWebparts.createStyleBlock({ children: borderRadius, title: "Border Radius", element: this.element, options: { sync: true } });
            this.paneStyle.append(borderRadiusBlock);

            let fonts = { fontSize: 'Font Size', fontWeight: 'Boldness', fontFamily: 'Font Style', color: 'Font Color' };
            let fontBlock = this.sharePoint.craterWebparts.createStyleBlock({ children: fonts, title: "fonts", element: this.element });
            this.paneStyle.append(fontBlock);

            let backgrounds = { backgroundColor: 'Background Color', backgroundSize: 'Background Size', backgroundRepeat: 'Background Repeat', backgroundImage: 'Background Image', backgroundPosition: 'Background Position', boxShadow: 'Box Shadow' };
            let backgroundBlock = this.sharePoint.craterWebparts.createStyleBlock({ children: backgrounds, title: "Backgrounds", element: this.element });
            this.paneStyle.append(backgroundBlock);

            let dimensions = { textAlign: 'Text Align', verticalAlign: 'Vertical Align', position: 'Position', visibility: 'Visibility', display: 'Display' };
            let dimensionBlock = this.sharePoint.craterWebparts.createStyleBlock({ children: dimensions, title: "Size", element: this.element });
            this.paneStyle.append(dimensionBlock);


            let height = { height: 'Height', minHeight: 'Minimium Height', maxHeight: 'Maximium Height' };
            let heightBlock = this.sharePoint.craterWebparts.createStyleBlock({ children: height, title: "Height", element: this.element });
            this.paneStyle.append(heightBlock);

            let width = { width: 'Width', minWidth: 'Minimium Width', maxWidth: 'Maximium Width' };
            let widthBlock = this.sharePoint.craterWebparts.createStyleBlock({ children: width, title: "Width", element: this.element });
            this.paneStyle.append(widthBlock);

            let display = { display: 'Display', gridTemplateColumns: 'Columns', gridTemplateRows: 'Rows' };
            let displayBlock = this.sharePoint.craterWebparts.createStyleBlock({ children: display, title: "Display", element: this.element });
            this.paneStyle.append(displayBlock);
        }

        this.editor.innerHTML = '';
        this.editor.append(this.paneStyle);

        this.paneStyle.findAll('.crater-style-attr').forEach(element => {
            element.onChanged(value => {
                let action = {};

                if (this.sharePoint.storage.pane.content[key].sync[element.dataset['styleSync']]) {
                    element.getParents('.crater-style-block').findAll('.crater-style-attr').forEach(styler => {
                        styler.value = value;
                        styler.setAttribute('value', value);
                        action[styler.dataset['action']] = value;
                        this.sharePoint.storage.pane.content[key].draft.dom.css(action);
                    });
                } else {
                    action[element.dataset['action']] = value;
                    this.sharePoint.storage.pane.content[key].draft.dom.css(action);
                }
            });
        });

        this.paneStyle.findAll('.crater-toggle-style-sync').forEach(element => {
            element.addEventListener('click', event => {
                this.sharePoint.storage.pane.content[key].sync[element.dataset['style']] = !this.sharePoint.storage.pane.content[key].sync[element.dataset['style']];
                element.src = this.sharePoint.storage.pane.content[key].sync[element.dataset['style']] ? this.sharePoint.images.sync : this.sharePoint.images.async;
            });
        });

        this.craterWebparts.sharePoint = this.sharePoint;

        let backgroundImageCell = this.paneStyle.find('#Background-Image-cell').parentNode;

        this.craterWebparts.uploadImage({ parent: backgroundImageCell }, (backgroundImage) => {
            this.sharePoint.storage.pane.content[key].draft.dom.setBackgroundImage(backgroundImage.src);
            backgroundImageCell.find('#Background-Image-cell').src = backgroundImage.src;
        });

        let backgroundColorCell = this.paneStyle.find('#Background-Color-cell').parentNode;
        this.craterWebparts.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.find('#Background-Color-cell') }, (backgroundColor) => {
            this.sharePoint.storage.pane.content[key].draft.dom.css({ backgroundColor });
            backgroundColorCell.find('#Background-Color-cell').value = backgroundColor;
            backgroundColorCell.find('#Background-Color-cell').setAttribute('value', backgroundColor);
        });

        let colorCell = this.paneStyle.find('#Font-Color-cell').parentNode;
        this.craterWebparts.pickColor({ parent: colorCell, cell: colorCell.find('#Font-Color-cell') }, (color) => {
            this.sharePoint.storage.pane.content[key].draft.dom.css({ color });
            colorCell.find('#Font-Color-cell').value = color;
            colorCell.find('#Font-Color-cell').setAttribute('value', color);
        });

        this.paneStyle.addEventListener('mutated', event => {
            this.sharePoint.storage.pane.content[key].draft.pane.styles = this.paneStyle.innerHTML;
            this.sharePoint.storage.pane.content[key].draft.html = this.sharePoint.storage.pane.content[key].draft.dom.outerHTML;
        });
    }

    private clearDraft(draft) {
        draft.pane.content = '';// clear draft content
        draft.pane.styles = '';// clear draft style
        draft.html = '';

        console.log('Draft cleared');
    }
}

export {
    PropertyPane
};