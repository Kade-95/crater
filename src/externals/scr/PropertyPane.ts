import { ElementModifier, func, CraterWebParts, ColorPicker } from "./index";

class PropertyPane {
    public sharePoint: any;
    public paneContent: any;
    public paneStyle: any;
    public elementModifier: any;
    public element: any;
    public editor: any;
    public craterWebparts: any;

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
    }

    public render(element): any {
        this.element = element;

        let key = this.element.dataset['key'];
        if (!func.isset(key)) {
            alert("This element cannot be editted");
            return;
        }

        if (!func.isset(this.sharePoint.properties.pane.content[key])) {
            this.sharePoint.properties.pane.content[key] = { content: '', styles: '', settings: {}, draft: { dom: '', html: '', pane: { content: '', styles: '' } } };
        }

        let editWindow = this.elementModifier.createElement({
            element: 'div', attributes: {
                class: 'crater-edit-window', style: { height: window.innerHeight + 'px', width: window.innerWidth + 'px' }
            }
        });

        let menus = this.elementModifier.menu({
            content: [
                { name: 'Content', owner: 'Content' },
                { name: 'Styles', owner: 'Styles' }
            ],
            padding: '1em 0em'
        });

        editWindow.appendChild(menus);

        this.editor = this.elementModifier.createElement({
            element: 'div', attributes: {
                class: 'crater-editor', style: {
                    height: `${8 * window.innerHeight / 10}px`,
                    width: `${9 * window.innerWidth / 10}px`,
                    marginTop: `${0.5 * window.innerHeight / 10}px`,
                    marginLeft: `${0.5 * window.innerWidth / 10}px`
                }
            }
        });

        editWindow.appendChild(this.editor);

        menus.querySelectorAll('.crater-menu-item').forEach(item => {
            item.addEventListener('click', event => {
                if (item.dataset.owner == 'Content') {
                    this.setUpContent(key);
                }
                else if (item.dataset.owner == 'Styles') {
                    this.setUpStyle(key);
                }
            });
        });

        editWindow.makeElement({
            element: 'div', attributes: { style: { position: 'absolute', bottom: '0px', marginBottom: '1.1em', right: '5%' } }, children: [
                this.elementModifier.createElement({
                    element: 'button', attributes: { id: 'crater-editor-save', class: 'button' }, text: 'Save'
                }),
                this.elementModifier.createElement({
                    element: 'button', attributes: { id: 'crater-editor-cancel', class: 'button' }, text: 'Cancel'
                })
            ]
        })
            .addEventListener('click', event => {
                if (event.target.id == 'crater-editor-save') {//save is clicked
                    this.sharePoint.properties.pane.content[key].styles = this.paneStyle.innerHTML;

                    if (func.isset(this.sharePoint.properties.pane.content[key].draft.dom.dataset.backgroundImage)) {
                        this.element.setBackgroundImage(this.sharePoint.properties.pane.content[key].draft.dom.dataset.backgroundImage);
                    }

                    this.sharePoint.saved = true;
                    this.sharePoint.savedWebPart = this.element;

                    console.log('Crater edition saved');
                } else if (event.target.id == 'crater-editor-cancel') {//keep draft and exit
                    this.clearDraft(this.sharePoint.properties.pane.content[key].draft);
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
        let type = this.element.dataset['type'];

        if (this.sharePoint.properties.pane.content[key].draft.html == '') {
            this.sharePoint.properties.pane.content[key].draft.dom = this.element.cloneNode(true);
            this.sharePoint.properties.pane.content[key].draft.html = this.sharePoint.properties.pane.content[key].draft.dom.outerHTML;
        } else {
            this.sharePoint.properties.pane.content[key].draft.dom = this.elementModifier.createElement(this.sharePoint.properties.pane.content[key].draft.html);
        }

        this.editor.innerHTML = '';
        this.editor.append(this.craterWebparts[type]({ action: 'setUpPaneContent', element: this.element, sharePoint: this.sharePoint }));

        this.craterWebparts[type]({ action: 'listenPaneContent', element: this.element, sharePoint: this.sharePoint });

        this.editor.querySelectorAll('.crater-color-picker').forEach(picker => {
            picker.remove();
        });
    }

    private setUpStyle(key): any {
        this.paneStyle = this.elementModifier.createElement({
            element: 'div',
            attributes: { class: 'crater-property-style' }
        }).monitor();

        if (this.sharePoint.properties.pane.content[key].draft.pane.styles != '') {
            this.paneStyle.innerHTML = this.sharePoint.properties.pane.content[key].draft.pane.styles;
        }
        else if (this.sharePoint.properties.pane.content[key].styles != '') {
            this.paneStyle.innerHTML = this.sharePoint.properties.pane.content[key].styles;
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

            let fonts = { fontSize: 'Font Size', fontWeight: 'Boldness', fontStyle: 'Font Style', fontFamily: 'Font Family', color: 'Font Color' };
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
        }

        this.editor.innerHTML = '';
        this.editor.append(this.paneStyle);

        this.paneStyle.querySelectorAll('.crater-style-attr').forEach(element => {
            element.onChanged(value => {
                let action = {};

                if (this.sharePoint.properties.pane.content[key].sync[element.dataset['styleSync']]) {
                    element.getParents('.crater-style-block').querySelectorAll('.crater-style-attr').forEach(styler => {
                        styler.value = value;
                        styler.setAttribute('value', value);
                        action[styler.dataset['action']] = value;
                        this.sharePoint.properties.pane.content[key].draft.dom.css(action);
                    });
                } else {
                    action[element.dataset['action']] = value;
                    this.sharePoint.properties.pane.content[key].draft.dom.css(action);
                }
            });
        });

        this.paneStyle.querySelectorAll('.crater-toggle-style-sync').forEach(element => {
            element.addEventListener('click', event => {
                this.sharePoint.properties.pane.content[key].sync[element.dataset['style']] = !this.sharePoint.properties.pane.content[key].sync[element.dataset['style']];
                element.src = this.sharePoint.properties.pane.content[key].sync[element.dataset['style']] ? this.sharePoint.images.sync : this.sharePoint.images.async;
            });
        });

        this.craterWebparts.sharePoint = this.sharePoint;

        let backgroundImageCell = this.paneStyle.querySelector('#Background-Image-cell').parentNode;

        this.craterWebparts.uploadImage({ parent: backgroundImageCell }, (backgroundImage) => {
            this.sharePoint.properties.pane.content[key].draft.dom.setBackgroundImage(backgroundImage.src);
            backgroundImageCell.querySelector('#Background-Image-cell').src = backgroundImage.src;
        });

        let backgroundColorCell = this.paneStyle.querySelector('#Background-Color-cell').parentNode;
        this.craterWebparts.pickColor({ parent: backgroundColorCell, cell: backgroundColorCell.querySelector('#Background-Color-cell') }, (backgroundColor) => {
            this.sharePoint.properties.pane.content[key].draft.dom.css({ backgroundColor });
            backgroundColorCell.querySelector('#Background-Color-cell').value = backgroundColor;
            backgroundColorCell.querySelector('#Background-Color-cell').setAttribute('value', backgroundColor);
        });

        let colorCell = this.paneStyle.querySelector('#Font-Color-cell').parentNode;
        this.craterWebparts.pickColor({ parent: colorCell, cell: colorCell.querySelector('#Font-Color-cell') }, (color) => {
            this.sharePoint.properties.pane.content[key].draft.dom.css({ color });
            colorCell.querySelector('#Font-Color-cell').value = color;
            colorCell.querySelector('#Font-Color-cell').setAttribute('value', color);
        });

        this.paneStyle.addEventListener('mutated', event => {
            this.sharePoint.properties.pane.content[key].draft.pane.styles = this.paneStyle.innerHTML;
            this.sharePoint.properties.pane.content[key].draft.html = this.sharePoint.properties.pane.content[key].draft.dom.outerHTML;
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