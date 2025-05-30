# LabelSP——SmartPointLabeler 🧷

> A minimal, open-source annotation tool for point-based object labeling.  
> Designed for datasets where object boundaries are vague, irregular, or point-centered (e.g., fish farms, lesions, small parts).

## ✨ Features

- 🔹 **Single or dual-point annotation** mode
- 📦 **Auto convert points to bounding boxes**
- 🧠 **Ideal for fuzzy, irregular or point-centered targets**
- 📄 **YOLO-friendly output format** with optional metadata
- 🎨 **Minimal GUI**, inspired by LabelMe/LabelImg
- 💻 Built with **Python + Tkinter**, no dependencies

## 🎬 Demo

![screenshot](./assets/demo_ui.gif)

> Users click on a single center point or two diagonal points. The tool auto-generates bounding box data based on predefined box size or logic.

## 🛠️ Installation

```bash
git clone https://github.com/your-username/SmartPointLabeler.git
cd SmartPointLabeler
python smart_label.py
