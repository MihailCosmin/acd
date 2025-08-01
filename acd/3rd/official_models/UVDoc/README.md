---
license: apache-2.0
library_name: PaddleOCR
language:
- en
- zh
pipeline_tag: image-to-text
tags:
- OCR
- PaddlePaddle
- PaddleOCR
- doc_img_unwarping
---

# UVDoc

## Introduction

The main purpose of text image correction is to carry out geometric transformation on the image to correct the document distortion, inclination, perspective deformation and other problems in the image, so that the subsequent text recognition can be more accurate.

| Model| CER | 
|  --- | --- | 
|UVDoc |  0.179 | 

**Note**: Test data set: docunet benchmark data set.

## Quick Start

### Installation

1. PaddlePaddle

Please refer to the following commands to install PaddlePaddle using pip:

```bash
# for CUDA11.8
python -m pip install paddlepaddle-gpu==3.0.0 -i https://www.paddlepaddle.org.cn/packages/stable/cu118/

# for CUDA12.6
python -m pip install paddlepaddle-gpu==3.0.0 -i https://www.paddlepaddle.org.cn/packages/stable/cu126/

# for CPU
python -m pip install paddlepaddle==3.0.0 -i https://www.paddlepaddle.org.cn/packages/stable/cpu/
```

For details about PaddlePaddle installation, please refer to the [PaddlePaddle official website](https://www.paddlepaddle.org.cn/en/install/quick).

2. PaddleOCR

Install the latest version of the PaddleOCR inference package from PyPI:

```bash
python -m pip install paddleocr
```

### Model Usage

You can quickly experience the functionality with a single command:

```bash
paddleocr text_image_unwarping --model_name UVDoc -i https://cdn-uploads.huggingface.co/production/uploads/63d7b8ee07cd1aa3c49a2026/SfMVKd0xnMII5KBDV6Mfz.jpeg
```

You can also integrate the model inference of the TextImageUnwarping module into your project. Before running the following code, please download the sample image to your local machine.

```python
from paddleocr import TextImageUnwarping

model = TextImageUnwarping(model_name="UVDoc")
output = model.predict("SfMVKd0xnMII5KBDV6Mfz.jpeg", batch_size=1)
for res in output:
    res.print()
    res.save_to_img(save_path="./output/")
    res.save_to_json(save_path="./output/res.json")
```

After running, the obtained result is as follows:

```json
{'res': {'input_path': 'doc_test.jpg', 'page_index': None, 'doctr_img': '...'}}
```

The visualized image is as follows:

![image/jpeg](https://cdn-uploads.huggingface.co/production/uploads/63d7b8ee07cd1aa3c49a2026/1405yNIYq_hA9VL3_8Itn.jpeg)

For details about usage command and descriptions of parameters, please refer to the [Document](https://paddlepaddle.github.io/PaddleOCR/latest/en/version3.x/module_usage/text_image_unwarping.html#iii-quick-integration).


### Pipeline Usage

The ability of a single model is limited. But the pipeline consists of several models can provide more capacity to resolve difficult problems in real-world scenarios.


#### PP-StructureV3

Layout analysis is a technique used to extract structured information from document images. PP-StructureV3 includes the following six modules:
* Layout Detection Module
* General OCR Sub-pipeline
* Document Image Preprocessing Sub-pipeline （Optional）
* Table Recognition Sub-pipeline （Optional）
* Seal Recognition Sub-pipeline （Optional）
* Formula Recognition Sub-pipeline （Optional）

You can quickly experience the PP-StructureV3 pipeline with a single command.

```bash
paddleocr pp_structurev3 --use_doc_unwarping True -i https://cdn-uploads.huggingface.co/production/uploads/63d7b8ee07cd1aa3c49a2026/KP10tiSZfAjMuwZUSLtRp.png
```

You can experience the inference of the pipeline with just a few lines of code. Taking the PP-StructureV3 pipeline as an example:

```python
from paddleocr import PPStructureV3

pipeline = PPStructureV3(use_doc_unwarping=True) # Use use_doc_unwarping to enable/disable document unwarping module
output = pipeline.predict("./KP10tiSZfAjMuwZUSLtRp.png")
for res in output:
    res.print() ## Print the structured prediction output
    res.save_to_json(save_path="output") ## Save the current image's structured result in JSON format
    res.save_to_markdown(save_path="output") ## Save the current image's result in Markdown format
```

For details about usage command and descriptions of parameters, please refer to the [Document](https://paddlepaddle.github.io/PaddleOCR/latest/en/version3.x/pipeline_usage/PP-StructureV3.html#2-quick-start).

## Links

[PaddleOCR Repo](https://github.com/paddlepaddle/paddleocr)

[PaddleOCR Documentation](https://paddlepaddle.github.io/PaddleOCR/latest/en/index.html)

