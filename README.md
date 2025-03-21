# Corrupted PowerPoint text extraction
- This is a small project that uses a couple of libraries to extract the text from large PowerPoint files that are corrupted. This is primarily for students or teachers trying to get through PowerPoint submissions where the file has gone bust and can't be openned.  
## Requirements 
### Basic Requirements 
- python 3.8+ 
- pip, this is the python package manager which is crucial for installs later on.
## How to install the program 
### Option 1: Command Line installation 
1. Clone the repository:
``` bash 
git clone https://github.com/RosslanKoulli/Corrupted-PowerPoint-text-extraction
cd Corrupted-PowerPoint-text-extraction
```
2. Crate a virtual environment: 
``` bash 
python -m venv venv
source venv/bin/activate  # On Windows, use: venv\Scripts\activate
```
3. Install the basic dependencies:
``` bash
pip install -r requirements.txt
```
4. For OCR functionality(optional): **On Ubuntu/Debian**
``` bash
sudo apt-get update
sudo apt-get install -y tesseract-ocr poppler-utils 
```
**On MAC OS:
``` bash
brew install tesseract poppler
```
**On Windows:
- Download [Tesseract]()
- Download [Poppler for Windows]()
- And both to your PATH

### Option 2: Using Docker(Recommended)
If you decide to use the docker option it will contain all the dependencies pre-installed on it:
``` bash 
# Pull the Docker Image
docker pull RosslanKoulli/Corrupted-PowerPoint-text-extraction

# Or build it locally 
docker build -t Corrupted-PowerPoint-text-extraction .
```
### What it can be used on

#### PowerPoint & PDF text Extractor 
It's primary use case is for the user to be able to extract text from large corrrupted PDF and PowerPoints. Image extraction is in the works but for now it can only extract text. This includes meta data. 
``` bash 
# Normal Installation
python pptxToTextConverter.py input.pptx --output extracted_text.txt

# Using docker
docker run --rm -v$(pwd):/data Corrupted-PowerPoint-text-extraction pptxToTextConverter.py /data/input.pptx --output /data/extracted_text.txt 
```
Options:
- `--output/-o`: Specify output file path
- `--verbose/-v`: Enable detailed logging 
- `--work-dir/w`: Sets directory for temporary files 


#### Text Preprocessor 
You can use this program for cleaning and processing the data into a PowerPoint file. 

``` bash
# Normal installation
python txtToPptxConverterProgram.py extracted_text.txt --output cleaned_text.txt --use-nlp

# Using Docker
docker run --rm -v $(pwd):/data Corrupted-PowerPoint-text-extraction txtToPptxConverterProgram.py /data/extracted_text.txt --output /data/cleaned_text.txt --use-nlp
```
Options:
- `--output/-o`: Specify output file path
- `--aggressive/-a`: User more aggressive cleaning 
- `--preserve-bullets/-b`: Preserve bullet points and numbering
- `--use-nlp/-n`: Use NLP for advanced content analysis
- `--verbose/-v`: Enable detailed logging 

### Example Workflow 
#### Normal Download
``` bash
# Extract text from PPTX/PDF
python pptxToTextConverter.py presentation.pptx --output raw_text.txt

#Clean the extracted text
python txtToPptxConverterProgram.py raw_text.txt --output cleaned_text.txt --use-nlp
```

#### Docker Container Method
First you have to build the docker image. The docker image will be built from the dockerFile(which is like the recipe) and the docker image is the virtual filesystem. The container is the running process on the hosts machine.

``` bash
docker build -t Corrupted-PowerPoint-text-extraction .
```
#### Runnign from docker 

``` bash
# Mount on your current directory to /data in the container
docker run --rm -v $(pwd):/data Corrupted-PowerPoint-text-extraction pptxToTextConverter.py /data/your_file.pptx --output /data/output.txt
```

