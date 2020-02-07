import { ColorPicker, func } from '.';

class ElementModifier {
    public sharepoint: any;
    constructor(params?) {
        //Get the sharepoint webpart
        if (func.isset(params)) this.sharepoint = params.sharePoint;
        //add the Element add-ons
        prepareFrameWork();
    }

    public dropDown(params) {
        let mDrop = this.createElement({
            element: 'span', attributes: params.attributes, children: [
                {
                    element: 'span', attributes: { id: 'drop-control', style: { display: 'grid', gridTemplateColumns: '1fr max-content' } }, children: [
                        { element: 'input', attributes: params.dataAttributes },
                        {
                            element: 'span', attributes: {
                                style: {
                                    cursor: 'pointer',
                                    transform: 'rotate(225deg)',
                                    margin: '.5em',
                                    borderTop: '2px solid black',
                                    borderLeft: '2px solid black',
                                    width: '.5em',
                                    height: '.5em',
                                    display: 'flex'
                                }, id: 'drop'
                            }
                        }
                    ]
                },
                {
                    element: 'span', attributes: {
                        id: 'drop-contents',
                        style: {
                            backgroundColor: 'white', display: 'none', height: '300px', overflow: 'auto'
                        }
                    }, children: params.contents
                }
            ]
        });

        for (let child of mDrop.querySelector('#drop-contents').childNodes) {
            child.classList.add('drop-content');
            if (!child.hasAttribute('value')) child.setAttribute('value', child.textContent);
            child.css({ cursor: 'pointer' });
            child.addEventListener('mouseenter', event => {
                child.css({ backgroundColor: 'lightgray' });
            });

            child.addEventListener('mouseleave', event => {
                child.cssRemove(['background-color']);
            });
        }

        mDrop.css({
            border: '2px solid black',
            display: 'grid',
        });

        let dropDownListener = () => {
            let drop = mDrop.querySelector('#drop');

            drop.addEventListener('click', event => {
                mDrop.querySelector('#drop-contents').toggle();
            });

            mDrop.querySelector('#drop-contents').addEventListener('click', event => {
                let clicked = event.target;
                if (!clicked.classList.contains('drop-content')) clicked = clicked.getParents('.drop-content');

                if (!func.isnull(clicked)) {
                    mDrop.querySelector('#drop-control').querySelector('input').value = clicked.getAttribute('value') || '';
                    mDrop.querySelector('#drop-control').querySelector('input').setAttribute('value', clicked.getAttribute('value') || '');
                }
            });
        };

        dropDownListener();

        return mDrop;
    }

    //Import image as string64
    public importImage(params, callBack) {
        params.attributes = func.isset(params.attributes) ? params.attributes : {};
        let upload = this.cell({
            element: 'input', name: params.name, dataAttributes: { type: 'file' }
        });

        //Upload as form

        let allImages = [];
        let contents = [];
        let style = {
            display: 'flex',
            justifyContent: 'space-around',
            overflow: 'hidden',
            margin: '0.5em 0em',
            padding: '0.2em 0em'
        };

        let getImages = () => {
            return new Promise((resolve, reject) => {

                for (let image of this.sharepoint.app.querySelectorAll('img')) {
                    // @ts-ignore
                    if (!allImages.includes(image.src)) {
                        allImages.push(image.src);
                        contents.push({
                            element: 'p', attributes: {
                                value: image.src,
                                class: 'single-image',
                                style
                            }, children: [
                                { element: 'img', attributes: { class: 'crater-icon', src: image.src } },
                                { element: 'a', text: image.src.slice(0, 20) }
                            ]
                        });
                    }
                }

                for (let storedImage in this.sharepoint.images) {
                    // @ts-ignore
                    if (!allImages.includes(this.sharepoint.images[storedImage]) && typeof this.sharepoint.images[storedImage] == 'string') {
                        allImages.push(this.sharepoint.images[storedImage]);
                        contents.push({
                            element: 'p', attributes: {
                                value: this.sharepoint.images[storedImage],
                                class: 'single-image',
                                style
                            }, children: [
                                { element: 'img', attributes: { class: 'crater-icon', src: this.sharepoint.images[storedImage] } },
                                { element: 'a', text: this.sharepoint.images[storedImage].slice(0, 20) }
                            ]
                        });
                    }
                }

                resolve(allImages);
            });
        };

        getImages().then(res => {
            let imageSelector = this.dropDown({
                attributes: { class: 'crater-image-selector' }, dataAttributes: { id: params.name, value: params.value || '' }, contents
            });

            let link = this.createElement({
                element: 'span', attributes: { style: { display: 'grid' } }, children: [
                    imageSelector,
                    { element: 'button', attributes: { id: `submit`, class: 'small-btn' }, text: 'Change' },
                ]
            });

            //close import window
            let closeButton = this.createElement({
                element: 'img', attributes: { class: 'crater-close small-btn', src: this.sharepoint.images.close }
            });

            closeButton.addEventListener('click', event => {
                params.parent.querySelectorAll('.upload-form').forEach(element => {
                    element.remove();
                });
            });

            params.attributes.style = (func.isset(params.attributes.style)) ? params.attributes.style : {};

            params.attributes.style.minHeight = '10px';
            params.attributes.style.position = 'relative';

            let selectionView = params.parent.makeElement({
                element: 'span', attributes: params.attributes, children: [
                    closeButton, link, upload
                ]
            });

            selectionView.css({ display: 'grid', gridGap: '0.5em', padding: '0.5em' });

            upload.querySelector(`#${params.name}-cell`).addEventListener('change', event => {
                this.imageToJson(upload.querySelector(`#${params.name}-cell`).files[0], (file) => {
                    callBack(file);
                });
            });

            link.querySelector(`#submit`).addEventListener('click', event => {
                let value = link.querySelector(`#${params.name}`).value;
                if (!func.isSpaceString(value)) callBack({ src: link.querySelector(`#${params.name}`).value });
            });

            imageSelector.ondblclick = (event) => {
                let clicked = event.target;
                if (!clicked.classList.contains('single-image')) clicked = clicked.getParents('.single-image');
                if (!func.isnull(clicked)) {
                    link.querySelector(`#submit`).click();
                }
            };
        });
    }

    //Convert image to object
    public imageToJson(file, callBack) {
        let fileReader = new FileReader();
        let myfile: any = {};
        fileReader.onload = (event: any) => {
            myfile.src = event.target.result;
            callBack(myfile);
        };

        myfile.size = file.size;
        myfile.type = file.type;
        fileReader.readAsDataURL(file);
    }

    public base64ToBlob(image, type = '', sliceSize = 512) {
        const byteChars = atob(image);
        const byteArrays = [];

        for (let offset = 0; offset < byteChars.length; offset += sliceSize) {
            const slice = byteChars.slice(offset, offset + sliceSize);

            const byteNumbers = new Array(slice.length);
            for (let i = 0; i < slice.length; i++) {
                byteNumbers[i] = slice.charCodeAt(i);
            }

            const byteArray = new Uint8Array(byteNumbers);
            byteArrays.push(byteArray);

        }

        const blob = new Blob(byteArrays, { type });
        return blob;
    }

    public b64ToBlob(url) {
        return new Promise((resolve, reject) => {
            fetch(url)
                .then(res => {
                    res.blob().then(image => {
                        resolve(image);
                    });
                });
        });
    }

    //Create a crater element
    public createElement(params) {

        let getElement = (param) => {
            var element: any;
            //if params is a HTML String
            if (typeof params == 'string') {
                let div = this.createElement({ element: 'div' });
                div.innerHTML = params;
                element = div.firstChild;
            }
            //if params is object
            else if (typeof params == 'object') {
                element = document.createElement(params.element);//generate the element
                if (func.isset(params.attributes)) {//set the attributes
                    for (var attr in params.attributes) {
                        if (attr == 'style') {//set the styles
                            element.css(params.attributes[attr]);
                        }
                        else element.setAttribute(attr, params.attributes[attr]);
                    }
                }
                if (func.isset(params.children)) {//add the children if set
                    for (var child of params.children) {
                        if (child instanceof Element) {
                            element.append(child);
                        } else if (typeof child === "object") {
                            element.makeElement(child);
                        }
                    }
                }
                if (func.isset(params.text)) element.textContent = params.text;//set the innerText
                if (func.isset(params.value)) element.value = params.value;//set the value
                if (func.isset(params.options)) {//add options if isset
                    for (var i of params.options) {
                        element.makeElement({ element: 'option', value: i, text: i, attachment: 'append' });
                    }
                }

                if (func.isset(params.currentContent)) {
                    element.value = params.currentContent;//if has content value is currentContent
                }
            }

            this.setCratetKey(element).then(key => {
                if (func.isset(params.state)) {
                    let owner = element.getParents(params.state.owner);
                    if (!func.isnull(owner)) {
                        owner.addState({ name: params.state.name, state: element });
                    }
                }
            });//Set the crater key and store it in craterDom
            return element;
        };

        if (Array.isArray(params)) {
            let elements = [];
            for (let param of params) {
                elements.push(getElement(param));
            }
            return elements;
        } else {
            let element = getElement(params);
            return element;
        }
    }

    public validateFormTextarea(element) {
        if (element.value == '') {
            return false;
        }
        return true;
    }

    public validateFormInput(element) {
        var type = element.getAttribute('type');
        var value = element.value;
        if (type == 'file' && value == '') {
            return false;
        }
        else if (type == 'text' || func.isnull(type)) {
            return !func.isSpaceString(value);
        }
        else if (type == 'date') {
            if (func.hasString(element.className, 'future')) {
                return func.isDate(value);
            } else {
                return func.isDateValid(value);
            }
        }
        else if (type == 'email') {
            return func.isEmail(value);
        }
        else if (type == 'number') {
            return func.isNumber(value);
        }
        else if (type == 'password') {
            return func.isPasswordValid(value);
        }
    }

    public validateFormSelect(element) {
        if (element.value == 0 || element.value == 'null') {
            return false;
        }
        return true;
    }

    public validateForm(form, nodeNames) {
        if (!func.isset(nodeNames)) nodeNames = 'INPUT, SELECT, TEXTAREA';
        var final = true,
            nodeName = '',
            elementValue = true,
            prototype = null;
        form.querySelectorAll(nodeNames).forEach(element => {
            nodeName = element.nodeName;
            if (!func.isnull(element.getParents('#content_prototype'))) {
                prototype = element.getParents('#content_prototype').id;
                if (prototype == 'content_prototype') {
                    elementValue = true;
                }
            }
            else if (nodeName == 'INPUT') {
                elementValue = this.validateFormInput(element);
            }
            else if (nodeName == 'SELECT') {
                elementValue = this.validateFormSelect(element);
            }
            else if (nodeName == 'TEXTAREA') {
                elementValue = this.validateFormTextarea(element);
            }
            if (final) final = elementValue;
        });
        return final;
    }

    //create form component
    public createForm(params) {
        var form = this.createElement({
            element: params.element || 'form', attributes: params.attributes, children: [
                { element: 'h3', attributes: { class: 'form-title' }, text: params.title },
                { element: 'section', attributes: { class: 'form-contents', style: { gridTemplateColumns: `repeat(${params.columns}, 1fr)` } } },
                { element: 'section', attributes: { class: 'form-buttons' } },
            ]
        });

        if (func.isset(params.parent)) params.parent.append(form);
        var note;
        let formContents = form.querySelector('.form-contents');

        for (let key in params.contents) {
            note = (func.isset(params.contents[key].note)) ? `(${params.contents[key].note})` : '';
            let block = formContents.makeElement({
                element: 'div', attributes: { class: 'form-single-content' }, children: [
                    { element: 'label', text: key.toLowerCase(), attributes: { class: 'form-label', for: key.toLowerCase() } }
                ]
            });

            block.makeElement(params.contents[key]).classList.add('form-data');
            if (func.isset(params.contents[key].note)) block.makeElement({ element: 'span', text: params.contents[key].note, attributes: { class: 'form-note' } });
        }

        for (let key in params.buttons) {
            form.querySelector('.form-buttons').makeElement(params.buttons[key]);
        }

        form.makeElement({ element: 'span', attributes: { class: 'form-error' }, state: { name: 'error', owner: `#${form.id}` } });

        return form;
    }

    //create table component
    public createTable(params) {
        //create the table element
        let table = this.createElement(
            { element: 'table', attributes: params.attributes }
        );

        table.classList.add('table');//add table to the class

        let tableHead = table.makeElement({//create the table-head
            element: 'thead', children: [
                { element: 'tr' }
            ]
        });

        let tableBody = table.makeElement({//create the table-body
            element: 'tbody'
        });

        let tableHeadRow = tableHead.querySelector('tr');//create the table-head-row

        let headers = [];//the headers
        for (let content of params.contents) {
            let tableBodyRow = tableBody.makeElement({//create a table-body-row
                element: 'tr', attributes: { class: params.rowClass }
            });

            for (let key in content) {
                if (headers.indexOf(key) == -1) {//check if key has been added to the headers
                    headers.push(key);
                    let tableHeadCell = tableHeadRow.makeElement({//create table-head-cell
                        element: 'th', text: key, attributes: { 'data-name': 'crater-table-data-' + key }
                    });
                }

                let tableBodyRowData = tableBodyRow.makeElement({//create table-body-cell
                    element: 'td', text: content[key], attributes: { 'data-name': 'crater-table-data-' + key }
                });
            }
        }

        return table;
    }

    public getTableData(table) {
        let header = table.querySelector('thead');
        let body = table.querySelector('tbody');

        let data = [];
        let heads = [];
        if (!func.isnull(header)) {
            for (let head of header.querySelectorAll('th')) {
                heads.push(head.textContent);
            }
        }

        let rows = body.querySelectorAll('tr');

        for (let row of rows) {
            let line = {};
            data.push(line);
            row = row.querySelectorAll('td');
            for (let i in row) {
                // @ts-ignore
                if (!isNaN(i)) {
                    line[heads[i] || i] = row[i].textContent;
                }
            }
        }

        return data;
    }

    public sortTable(table, by?, direction?) {
        let data = this.getTableData(table);

        data.sort((a, b) => {
            a = a[by];
            b = b[by];

            if (func.isNumber(a) && func.isNumber(b)) {
                a = a / 1;
                b = b / 1;
            }

            if (direction > -1) {
                return a > b ? 1 : -1;
            }
            else {
                return a > b ? -1 : 1;
            }
        });
        return data;
    }

    //create options component
    public options(params) {
        //create the options element
        var options = this.createElement({ element: 'span', attributes: params.attributes });

        for (var i of params.options) {
            //append all the options
            options.append(
                this.createElement({ element: 'img', attributes: { src: `./../images/${i}.png`, alt: i, class: 'option ' + i, title: i } })
            );
        }

        //toggle the options
        params.parent.toggleChild(options);

        options.addEventListener('click', event => {

        });

        options.querySelectorAll('.option').forEach(element => {

        });

        return options;
    }

    //create cell component
    public cell(params) {
        //set the cell-data id
        var id = func.stringReplace(params.name, ' ', '-') + '-cell';

        //create the cell label
        var label = this.createElement({ element: 'label', attributes: { class: 'cell-label' }, text: params.name });

        //cell attributes
        params.attributes = (func.isset(params.attributes)) ? params.attributes : {};

        //cell data attributes
        params.dataAttributes = (func.isset(params.dataAttributes)) ? params.dataAttributes : {};
        params.dataAttributes.id = id;

        var components;

        //set the properties of cell data
        if (params.element == 'select') {//check if cell data is in select element
            components = {
                element: params.element, attributes: params.dataAttributes, children: [
                    { element: 'option', attributes: { disabled: '', selected: '' }, text: `Select ${params.name}`, value: '' }//set the default option
                ]
            };
        }
        else {
            components = { element: params.element, attributes: params.dataAttributes, text: params.value };
        }

        if (func.isset(params.value)) components.attributes.value = params.value;
        if (func.isset(params.options)) components.options = params.options;

        var data = this.createElement(components);//create the cell-data
        data.classList.add('cell-data');

        if (func.isset(params.value)) data.value = params.value;

        //create cell element
        var cell = this.createElement({ element: 'span', attributes: params.attributes, children: [label, data] });

        cell.classList.add('cell');

        if (func.isset(params.text)) data.text = params.text;

        if (func.isset(params.list)) {
            cell.makeElement({
                element: 'datalist', attributes: { id: `${id}-list` }, options: params.list.sort()
            });

            data.setAttribute('list', `${id}-list`);
        }

        if (func.isset(params.text)) data.textContent = params.text;
        return cell;
    }

    //create menu component
    public menu(params) {
        //create the menu element
        let menus = this.createElement({
            element: 'ul', attributes: { class: 'crater-menu' }
        });

        //add the menu children and set the width
        for (let menu of params.content) {
            menus.makeElement({
                element: 'li', attributes: { id: `${menu.owner.toLowerCase()}-menu-item`, class: 'crater-menu-item', 'data-owner': menu.owner }, children: [
                    { element: 'img', attributes: { class: 'crater-menu-item-icon', src: menu.icon || '' } },
                    { element: 'a', attributes: { class: 'crater-menu-item-text' }, text: menu.name }
                ]
            });
        }
        menus.css({ gridTemplateColumns: `repeat(${params.content.length}, 1fr)` });
        return menus;
    }

    public message(params) {
        var me = this.createElement({
            element: 'span', attributes: { class: 'alert' }, children: [
                func.isset(params.link) ?
                    this.createElement({ element: 'a', text: params.text, attributes: { class: 'text', href: params.link } })
                    :
                    this.createElement({ element: 'a', text: params.text, attributes: { class: 'text' } }),
                ,
                this.createElement({ element: 'span', attributes: { class: 'close' } })
            ]
        });

        if (func.isset(params.temp)) {
            var time = setTimeout(() => {
                me.remove();
                clearTimeout(time);
            }, (params.temp != '') ? params.time * 1000 : 5000);
        }

        me.querySelector('.close').addEventListener('click', event => {
            me.remove();
        });

        params.parent.querySelector('#notification-block').append(me);
    }

    public setCratetKey(element) {
        return new Promise((resolve, reject) => {
            let key = '';
            let found = false;
            if (!func.isset(window['craterdom'])) window['craterdom'] = {};
            if (!element.hasAttribute('domKey')) {
                do {
                    key = func.generateRandom(32);
                    found = func.isset(window['craterdom'][key]);
                } while (found);

                element.dataset.craterKey = key;
                window['craterdom'][key] = this;
            }
            resolve(key);
        });
    }

    public popUp(params) {
        let popUp = this.createElement({
            element: 'div', attributes: { class: 'crater-pop-up' }, children: [
                {
                    element: 'div', attributes: { class: 'crater-pop-up-window' }, children: [
                        {
                            element: 'span', attributes: { class: 'crater-pop-up-controls' }, children: [
                                { element: 'img', attributes: { src: params.maximize, class: 'crater-pop-up-controls-button', id: 'crater-pop-up-toggle' } },
                                { element: 'img', attributes: { src: params.close, class: 'crater-pop-up-controls-button', id: 'crater-pop-up-close' } }
                            ]
                        },
                        {
                            element: 'iframe', attributes: { class: 'crater-pop-up-content', src: params.source }
                        }
                    ]
                }
            ]
        });

        let window = popUp.querySelector('.crater-pop-up-window');

        popUp.addEventListener('click', event => {
            if (event.target.id == 'crater-pop-up-close') {
                popUp.remove();
            }
            else if (!event.target.classList.contains('crater-pop-up-window') && func.isnull(event.target.getParents('.crater-pop-up-window'))) {
                popUp.remove();
            }
            else if (event.target.id == 'crater-pop-up-toggle') {
                window.classList.toggle('full-window');
                if (window.classList.contains('full-window')) {
                    window.css({ width: '100%', height: '100vh' });
                    event.target.src = params.minimize;
                }
                else {
                    window.cssRemove(['width', 'height']);
                    event.target.src = params.maximize;
                }
            }
        });

        return popUp;
    }

    public choose(params) {
        let chooseWindow = this.createElement({
            element: 'span', attributes: { class: 'crater-choose' }, children: [
                { element: 'p', attributes: { class: 'crater-choose-note' }, text: params.note },
                { element: 'span', attributes: { class: 'crater-choose-control' } },
                { element: 'button', attributes: { id: 'crater-choose-close', class: 'btn' }, text: 'Close' }
            ]
        });

        let chooseControl = chooseWindow.querySelector('.crater-choose-control');

        chooseWindow.querySelector('#crater-choose-close').addEventListener('click', event=>{
            chooseWindow.remove();
        });

        for (let option of params.options) {
            chooseControl.makeElement({
                element: 'button', attributes: { class: 'btn choose-option' }, text: option
            });
        }

        return {
            display: chooseWindow, choice: new Promise((resolve, reject) => {
                chooseControl.addEventListener('click', event => {
                    if (event.target.classList.contains('choose-option')) {
                        resolve(event.target.textContent);
                        chooseWindow.remove();
                    }
                });
            })
        };
    }
}

function prepareFrameWork(): void {
    //Framework with JsDom

    NodeList.prototype['indexOf'] = function (element) {
        for (let i in this) {
            if (this[i] == element) return i;
        }
        return -1;
    };

    Element.prototype['getMostOccurredNode'] = function () {
        let children = this.childNodes;
        let frequency = {};
        for (let child of children) {
            let nodeName = child.nodeName;
            if (!func.isset(frequency[nodeName])) frequency[nodeName] = 0;
            else frequency[nodeName]++;
        }

        let mostOccurred;
        let occurrance = 0;
        for (let name in frequency) {
            if (frequency[name] > occurrance) {
                occurrance = frequency[name];
                mostOccurred = name;
            }
        }

        return mostOccurred;
    };

    Element.prototype['indexOf'] = function (element) {
        for (let i in Array(this.children)) {
            if (this.children[i] == element) return i;
        }
        return -1;
    };

    Element.prototype['setBackgroundImage'] = function (url) {
        if (url.indexOf('data:image/png;base64,') != -1) {
            let data = url.split('data:image/png;base64,')[1];
            this.css({ backgroundImage: "url('data:image/png;base64, " + data + "')" });
        } else {
            this.css({ backgroundImage: "url('" + url + "')" });
        }
        this.dataset.backgroundImage = url;
        return this;
    };

    Element.prototype['addState'] = function (params) {
        if (func.isset(params.state.dataset.domKey)) {
            this.dataset[params.name] = params.state.dataset.domKey;
            return true;
        }
        return false;
    };

    Element.prototype['setState'] = function (params) {
        let element = window['craterDom'][this.dataset[params.name]];

        if (func.isset(params.attributes)) {//set the attributes
            for (var attr in params.attributes) {
                if (attr == 'style') {//set the styles
                    element.css(params.attributes[attr]);
                }
                else element.setAttribute(attr, params.attributes[attr]);
            }
        }

        if (func.isset(params.children)) {//add the children if set
            for (var child of params.children) {
                if (child instanceof Element) {
                    element.append(child);
                } else if (typeof child === "object") {
                    element.makeElement(child);
                }
            }
        }

        if (func.isset(params.text)) element.textContent = params.text;//set the innerText

        if (func.isset(params.value)) element.value = params.value;//set the value

        if (func.isset(params.options)) {//add options if isset
            for (var i of params.options) {
                element.makeElement({ element: 'option', value: i, text: i, attachment: 'append' });
            }
        }
    };

    Element.prototype['setAttributes'] = function (attributes) {
        for (let i in attributes) {
            if (i == 'style') {
                this.css(attributes[i]);
            }
            else {
                this.setAttribute(i, attributes[i]);
            }
        }
    };

    Element.prototype['before'] = function (element) {
        this.parentNode.appendBefore(element, this);
        return this;
    };

    Element.prototype['after'] = function (element) {
        this.parentNode.appendAfter(element, this);
        return this;
    };

    Element.prototype['appendBefore'] = function (newSibling, sibling) {
        this.insertBefore(newSibling, sibling);
        return this;
    };

    Element.prototype['appendAfter'] = function (newSibling, sibling) {
        this.insertBefore(newSibling, sibling.nextSibling);
        return this;
    };

    Element.prototype['stopMonitor'] = function () {
        if (this.observe) this.observer.disconnect();
        return this;
    };

    Element.prototype['monitor'] = function (config = { attributes: true, childList: true, subtree: true }) {
        this.observer = new MutationObserver((mutationList, observer) => {
            if (mutationList.length) this.dispatchEvent(new CustomEvent('mutated'));
        });

        this.observer.observe(this, config);
        return this;
    };

    Element.prototype['render'] = function (params) {
        this.innerHTML = '';
        this.makeElement(params);
    };

    Element.prototype['getCssProperties'] = function (property) {
        var styleSheets: any = Array(document.styleSheets),//get all the css styles files and rules
            cssRules,
            id = this.id,
            nodeName = this.nodeName,
            classList = Array(this.classList),
            properties = {},
            selectorText;

        for (let i in classList) classList[i] = `.${classList[i]}`;//turn each class to css class format [.class]

        for (let i = 0; i < styleSheets.length; i++) {//loop through all the css rules in document/app
            cssRules = styleSheets[i].cssRules;
            for (var j = 0; j < cssRules.length; j++) {
                selectorText = cssRules[j].selectorText; //for each selector text check if element has it as a css property
                if (selectorText == `#${id}` || selectorText == nodeName || classList.indexOf(selectorText) != -1) {
                    properties[selectorText] = cssRules[j].style[property];//then add to the css property of the element
                }
            }
        }

        //if element has property add it to css property
        if (func.isset(this.css()[property])) properties['style'] = this.css()[property];

        return properties;//return property as json
    };

    Element.prototype['hasCssProperty'] = function (property) {
        var properties = this.getCssProperties(property); //get elements css properties
        for (var i in properties) {//loop through json object
            if (func.isset(properties[i]) && properties[i] != '') {
                return true;// if property is found return true
            }
        }
        return false;
    };

    Element.prototype['cssPropertyValue'] = function (property) {
        //check for the value of a property of an element
        var properties = this.getCssProperties(property),
            id = this.id,
            classList = Array(this.classList);
        if (func.isset(properties['style']) && properties['style'] != '') return properties['style'];//check if style rule has the propert and return it's value
        if (func.isset(id) && func.isset(properties[`#${id}`]) && properties[`#${id}`] != '') return properties[`#${id}`];//check if element id rule has the propert and return it's value
        for (var i of classList) {//check if any class rule has the propert and return it's value
            if (func.isset(properties[`.${i}`]) && properties[`.${i}`] != '') return properties[`.${i}`];
        }
        //check if node rule has the propert and return it's value
        if (func.isset(properties[this.nodeName]) && properties[this.nodeName] != '') return properties[this.nodeName];
        return '';
    };

    Element.prototype['slide'] = function (params) {
        //move or slide element around the page
        var position = this.position();//get elements current postion

        if (this.cssPropertyValue('position') == '') {
            //if element does not have position property set position to relative
            this.css({ position: 'relative' });
        }

        //get element's position data (if not set assign empty array)
        var previousPositions = (this.dataset.position) ? JSON.parse(this.dataset.position) : [];
        previousPositions.push(position);// append element position to data position
        this.setAttribute('data-position', JSON.stringify(previousPositions));
        // default distance is 50px
        let distance = 50,
            i = parseInt(this.css().left) | 0,
            limit = 0,
            work = me => { return ++me; };

        if (func.isset(params)) {
            params.slide = 'left';
            distance = (func.isset(params.distance)) ? parseInt(params.distance) : distance;

            if (params.direction == 'left') {
                limit = position.left - distance;
                work = me => { return --me; };
            }
            else if (params.direction == 'up') {
                limit = position.top + distance;
                work = me => { return --me; };
                params.slide = 'top';
                i = parseInt(this.css().top) | 0;
            }
            else if (params.direction == 'down') {
                limit = position.top - distance;
                work = me => { return ++me; };
                params.slide = 'top';
                i = parseInt(this.css().top) | 0;
            }
            else {
                limit = position.left + distance;
                params.slide = 'left';
            }
        }
        else {
            params = {};
            limit = position.left + distance;
            params.slide = 'left';
        }

        var sliding = setInterval(() => {
            if (params.slide == 'top') {
                this.css({ top: `${i = work(i)}px` });
            }
            else {
                this.css({ left: `${i = work(i)}px` });
            }

            distance--;
            if (distance == 0) {
                // this.cssRemove('position')
                if (func.isset(params.finish)) {
                    params.finish({ self: this, animation: sliding });
                } else {
                    clearInterval(sliding);
                }
            }
        }, 10);
    };

    Element.prototype['slideIn'] = function (params) {
        this.css({ visibility: 'visible' });

        var position = this.getBoundingClientRect(),
            previousPositions = (this.dataset.position) ? JSON.parse(this.dataset.position) : [],
            distance = 50;

        previousPositions = (previousPositions.length > 0) ? previousPositions[previousPositions.length - 1] : position;

        params = (func.isset(params)) ? params : {};

        // if (params.direction == 'down') {
        //     distance = (window.innerHeight | window.outerHeight) + position.height - position.bottom;
        // }
        // else if (params.direction == 'up') {
        //     distance = position.bottom;
        // }
        // else if (params.direction == 'left') {
        //     distance = position.right;
        // }
        params.distance = distance;
        params.finish = result => {
            result.self.css({ visibility: 'hidden' });
            clearInterval(result.animation);
        };
        this.slide(params);
    };

    Element.prototype['slideOut'] = function (params) {
        var position = this.getBoundingClientRect(),
            distance = (window.innerWidth | window.outerWidth) + position.width - position.right;

        params = (func.isset(params)) ? params : {};

        if (params.direction == 'down') {
            distance = (window.innerHeight | window.outerHeight) + position.height - position.bottom;
        }
        else if (params.direction == 'up') {
            distance = position.bottom;
        }
        else if (params.direction == 'left') {
            distance = position.right;
        }
        params.distance = distance;
        params.finish = result => {
            result.self.css({ visibility: 'hidden' });
            clearInterval(result.animation);
        };
        this.slide(params);
    };

    Element.prototype['css'] = function (params) {
        // set css style of element using json        
        if (func.isset(params)) {
            Object.keys(params).map((styleKey) => {
                this.style[styleKey] = params[styleKey];
            });
        }

        let css = this.style.cssText,
            style = {},
            key,
            value;

        if (css != '') {
            css = css.split('; ');
            let pair;
            for (let i of css) {
                pair = func.trem(i);
                key = func.jsStyleName(pair.split(':')[0]);
                value = func.stringReplace(pair.split(':').pop(), ';', '');
                if (key != '') {
                    style[key] = func.trem(value);
                }
            }
        }

        return style;
    };

    Element.prototype['cssRemove'] = function (elements) {
        //remove a group of properties from elements style
        if (Array.isArray(elements)) {
            for (var i of elements) {
                this.style.removeProperty(i);
            }
        }
        else {
            this.style.removeProperty(elements);
        }
        return this.css();
    };

    Element.prototype['toggleChild'] = function (child) {
        //Add child if element does not have a child else remove the child form the element
        var name, _classes, id, found = false;
        Array(this.children).forEach(node => {
            name = node.nodeName;
            _classes = node.classList;
            id = node.id;
            if (name == child.nodeName && id == child.id && _classes.toString() == child.classList.toString()) {
                node.remove();
                found = true;
            }
        });
        if (!found) this.append(child);
    };

    Element.prototype['removeClass'] = function (_class) {
        this.classList.remove(_class);
        return this;
    };

    Element.prototype['hasClassList'] = function (classList) {
        var classes = this.classList.toString().split(',');
        classList = classList.toString('');
    };

    Element.prototype['addClass'] = function (_class) {
        this.classList.add(_class);
        return this;
    };

    Element.prototype['toggleClass'] = function (_class) {
        (this.classList.contains(_class)) ? this.classList.remove(_class) : this.classList.add(_class);
        return this;
    };

    Element.prototype['position'] = function (params) {
        if (func.isset(params)) {
            Object.keys(params).map(key => {
                params[key] = (new String(params[key]).slice(params[key].length - 2) == 'px') ? params[key] : `${params[key]}px`;
            });
            this.css(params);
        }

        return this.getBoundingClientRect();
    };

    Element.prototype['hasClass'] = function (_class) {
        var classes = this.classList.toString().split(',');
        return (classes.indexOf(_class) != -1);
    };

    Element.prototype['makeElement'] = function (params) {
        let element: any;
        if (params.element instanceof Element) {
            element = params.element;
        } else {
            let elementModifier = new ElementModifier;
            element = elementModifier.createElement(params);
        }

        if (Array.isArray(params)) {
            for (let i in params) {
                if (!func.isset(params[i].attachment)) params[i].attachment = 'append';
                this[params[i].attachment](element[i]);
            }
        } else {
            if (!func.isset(params.attachment)) params.attachment = 'append';
            this[params.attachment](element);
        }

        return element;
    };

    Element.prototype['onChanged'] = function (callBack) {
        let value = this.getAttribute('value');

        let updateMe = (event) => {
            if (event.target.nodeName == 'INPUT') {
                if (event.target.type == 'date') {
                    if (func.isDate(this.value)) this.setAttribute('value', this.value);
                }
                else if (event.target.type == 'time') {
                    if (func.isTimeValid(this.value)) this.setAttribute('value', this.value);
                }
                else {
                    this.setAttribute('value', this.value);
                }
            } else if (event.target.nodeName == 'SELECT') {
                for (let i = 0; i < event.target.options.length; i++) {
                    if (i == event.target.selectedIndex) {
                        event.target.options[i].setAttribute('selected', true);
                    } else {
                        event.target.options[i].removeAttribute('selected');
                    }
                }
            }

            if (func.isset(callBack)) {
                callBack(this.value);
            }
        };

        this.addEventListener('keyup', (event) => {
            updateMe(event);
        });

        this.addEventListener('change', (event) => {
            updateMe(event);
        });

        this.addEventListener('mutated', event => {
            updateMe(event);
        });
    };

    Element.prototype['getParents'] = function (name) {
        var attribute = name.slice(0, 1);
        var parent = this.parentNode;
        if (attribute == '.') {
            while (parent) {
                if (func.isset(parent.classList) && parent.classList.contains(name.slice(1))) {
                    break;
                }
                parent = parent.parentNode;
            }
        }
        else if (attribute == '#') {
            while (parent) {
                if (func.isset(parent.id) && parent.id == name.slice(1)) {
                    break;
                }
                parent = parent.parentNode;
            }
        }
        else {
            while (parent) {
                if (func.isset(parent.nodeName) && parent.nodeName.toLowerCase() == name.toLowerCase()) {
                    break;
                }
                else if (func.isset(parent.hasAttribute) && parent.hasAttribute(name)) {
                    break;
                }
                parent = parent.parentNode;
            }
        }
        return parent;
    };

    Element.prototype['hide'] = function (params) {
        this.css({ display: 'none' });
        return this;
    };

    Element.prototype['show'] = function (params) {
        this.cssRemove(['display']);
        return this;
    };

    Element.prototype['fadeIn'] = function (params) {
        var opacity: number = 0;
        this.style.opacity = opacity;
        var duration = (!func.isset(params) || !func.isset(params.duration)) ? 1000 : params.duration;
        var speed = (!func.isset(params) || !func.isset(params.duration)) ? duration / 1000 : params.speed;

        if (this.style.display == 'none') this.style.display = params.display;
        if (this.style.visibility == 'hidden') this.style.visibility = 'visible';

        var fading = setInterval(() => {
            opacity++;
            if (opacity == duration) {
                if (func.isset(params) && func.isset(params.reflect)) params.reflect(this);
                if (func.isset(params) && func.isset(params.finish)) params.finish(this);
                clearInterval(fading);
            }
            this.style.opacity = opacity / duration;
        }, speed);
    };

    Element.prototype['fadeOut'] = function (params) {
        var duration = (!func.isset(params) || !func.isset(params.duration)) ? 100 : params.duration;
        var speed = (!func.isset(params) || !func.isset(params.duration)) ? duration / 1000 : params.speed;
        var opacity: number = duration;

        var fading = setInterval(() => {
            opacity--;

            if (opacity == 0) {
                this.style.display = 'none';
                this.style.visibility = 'hidden';
                if (func.isset(params) && func.isset(params.reflect)) params.reflect(this);
                if (func.isset(params) && func.isset(params.finish)) params.finish(this);
                clearInterval(fading);
            }
            this.style.opacity = opacity / duration;
        }, speed);
        return this;
    };

    Element.prototype['fadeToggle'] = function (params) {
        if (this.style.display == 'none' || this.style.visibility == 'hidden') {
            this.fadeIn(params);
        }
        else {
            this.fadeOut(params);
        }
    };

    Element.prototype['toggle'] = function (params) {
        if (this.style.display == 'none') this.style.display = 'inline-block';
        else if (this.style.display != 'none') this.style.display = 'none';

        if (this.style.visibility == 'hidden') this.style.display = 'visible';
        else if (this.style.visibility != 'hidden') this.style.display = 'hidden';
    };

    Element.prototype['replaceWith'] = function (element: HTMLElement) {
        this.before(element);
        this.style.visibility = 'hidden';
        element.style.visibility = 'visible';
        var position = this.getBoundingClientRect();
        element.style.top = position.top + 'px';
        element.style.left = position.left + 'px';
    };

    Element.prototype['removeChildren'] = function (params) {
        this.childNodes.forEach(node => {
            if (func.isset(params)) {
                if (!((func.isset(params.name) && params.name == node.nodeName) || func.isset(params.class) && func.hasArrayElement(Array(node.className), params.class.split(' ')) || (func.isset(params.id) && params.id == node.id))) {
                    node.remove();
                }
            } else {
                node.remove();
            }
        });
    };

    Element.prototype['toggleChildren'] = function (params) {
        Array(this.children).forEach(node => {
            if (func.isset(params)) {
                if (!((func.isset(params.name) && params.name == node.nodeName) || func.isset(params.class) && func.hasArrayElement(Array(node.classList), params.class.split(' ')) || (func.isset(params.id) && params.id == node.id))) {
                    node.toggle();
                }
            } else {
                node.toggle();
            }
        });
    };
}

export { ElementModifier };