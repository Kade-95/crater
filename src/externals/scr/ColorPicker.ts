import { ElementModifier, func } from '.';
require('./../styles/colorpicker.css');

class ColorPicker {
    private target: any;
    private width: any;
    private height: any;
    private context: any;
    private pickerCircle: any;
    private interval: any;
    private pickedColor: any;
    public canvas: any;

    private elementModifier: any = new ElementModifier();

    constructor(params) {
        this.canvas = this.generateCanvas();
        this.target = this.canvas.querySelector('#color-picker');
        this.width = params.width;
        this.height = params.height;
        this.target.width = this.width;
        this.target.height = this.height;

        //the context
        this.context = this.target.getContext('2d');
        this.canvas.querySelector('.crater-picked-color-value').innerText = params.color;

        //Circle color
        this.pickerCircle = { x: 10, y: 10, width: 7, height: 7 };

        this.listen();
    }

    private generateCanvas() {
        let canvasContainer = this.elementModifier.createElement({
            element: 'div', attributes: { class: 'crater-color-picker' }, children: [
                { element: 'canvas', attributes: { id: 'color-picker', class: 'crater-canvas' } },
                {
                    element: 'div', attributes: { class: 'crater-color-picker-result' }, children: [
                        { element: 'span', attributes: { class: 'crater-picked-color' } },
                        {
                            element: 'span', attributes: { class: 'crater-picked-color-value' }
                        },
                    ]
                },
                { element: 'div', text: 'Close', attributes: { class: 'crater-close-color-picker' } }
            ]
        });

        return canvasContainer;
    }

    private build() {
        let gradient = this.context.createLinearGradient(0, 0, this.width, 0);

        //color stops
        gradient.addColorStop(0, "rgb(255, 0, 0)");
        gradient.addColorStop(0.15, "rgb(255, 0, 255)");
        gradient.addColorStop(0.33, "rgb(0, 0, 255)");
        gradient.addColorStop(0.49, "rgb(0, 255, 255)");
        gradient.addColorStop(0.67, "rgb(0, 255, 0)");
        gradient.addColorStop(0.87, "rgb(255, 255, 0)");
        gradient.addColorStop(1, "rgb(255, 0, 0)");

        this.context.fillStyle = gradient;
        this.context.fillRect(0, 0, this.width, this.height);

        //add black and white stops
        gradient = this.context.createLinearGradient(0, 0, 0, this.height);
        gradient.addColorStop(0, "rgba(255, 255, 255, 1)");
        gradient.addColorStop(0.5, "rgba(255, 255, 255, 0)");
        gradient.addColorStop(0.5, "rgba(0, 0, 0, 0)");
        gradient.addColorStop(1, "rgba(0, 0, 0, 1)");

        this.context.fillStyle = gradient;
        this.context.fillRect(0, 0, this.width, this.height);

        //circle picker
        this.context.beginPath();
        this.context.arc(this.pickerCircle.x, this.pickerCircle.y, this.pickerCircle.width, 0, Math.PI * 2);
        this.context.strokeStyle = 'black';
        this.context.stroke();
        this.context.closePath();
    }

    private listen() {
        let isMouseDown = false;

        const onMouseDown = (event) => {
            let currentX = event.clientX - this.target.getBoundingClientRect().left;
            let currentY = event.clientY - this.target.getBoundingClientRect().top;

            //is mouse in color picker
            isMouseDown = (currentX > 0 && currentX < this.target.getBoundingClientRect().width && currentY > 0 && currentY < this.target.getBoundingClientRect().height);
        };

        const onMouseMove = (event) => {
            if (isMouseDown) {
                this.pickerCircle.x = event.clientX - this.target.getBoundingClientRect().left;
                this.pickerCircle.y = event.clientY - this.target.getBoundingClientRect().top;

                let picked = this.getPickedColor();

                this.pickedColor = `rgb(${picked.r}, ${picked.g}, ${picked.b})`;
                this.target.dispatchEvent(new CustomEvent('colorChanged'));

                this.canvas.querySelector('.crater-picked-color').css({ backgroundColor: this.pickedColor });

                this.canvas.querySelector('.crater-picked-color-value').innerText = this.pickedColor;
            }
        };

        const onMouseUp = (event) => {
            isMouseDown = false;
        };

        //Register
        this.target.addEventListener("mousedown", onMouseDown);
        this.target.addEventListener("mousemove", onMouseMove);
        document.addEventListener("mouseup", onMouseUp);

        this.canvas.querySelector('.crater-close-color-picker').addEventListener('click', event => {
            this.dispose();
        });
    }

    public onChanged = (callBack) => {
        this.target.addEventListener('colorChanged', event => {
            callBack(this.pickedColor);
        });
    }

    public getPickedColor() {
        let imageData = this.context.getImageData(this.pickerCircle.x, this.pickerCircle.y, 1, 1);
        return { r: imageData.data[0], g: imageData.data[1], b: imageData.data[2] };
    }

    public draw(speed) {
        this.interval = setInterval(() => this.build(), speed);
    }

    public dispose() {
        clearInterval(this.interval);
        this.canvas.remove();
    }

    private rgbToHex(color) {
        let hex = color.match(/\d+/g).map(x => {
            return parseInt(x).toString(16);
        });
        return '#' + hex.join('').toUpperCase();
    }
}

export { ColorPicker };