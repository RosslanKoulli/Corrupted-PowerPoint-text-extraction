.PHONY: build and extract preprocess clean help

# Deafult target

help:
  @echo "PDF/ PPTX Text extractiona dn precoessing tool"
  @echo ""
  @echo "Available commands"
  @echo "  make build			- build the Docker image"
  @echo "  make extract 		- Extract text from a PPTX file (set PPTX=path/to/file.pptx)"
  @echo "  make preprocess 		- precoess text file (set TXT=path/to/file.txt)"
  @echo "  make clean 			- Remove temporary files"
  @echo	""
  @echo "Examples:"
  @echo " Make extract PPTX= ,my_presentation.pptx"
  @echo " Make preprocess TXT=extracted_text.txt USE_NLP=1"

#Build docker image
build:
  docker-compose build

#extract text from PPTX

extract:
ifndef PPTX
  @echo "Error: PPTx file not specififed. Use PPTX=path/to/file.pptx"
  @exit 1
endif
  @OUTPUT=$$(basename $(PPTX) .pptx).txt; \
  docker-compose run --rm pptx-text-tools pptxToTextConverter.py /data/$(PPTX) --output /data/$$OUTPUT;	\
  echo "Extracted text saved to $$OUTPUT"

#Preprocess text file
preprocess:
ifndef TXT
	@echo "Error: Text file not specified. Use TXT=path/to/file.txt"
	@exit 1
endif
	@OUTPUT=$$(basename $(TXT) .txt)_cleaned.txt; \
	CMD="txtToPptxConverterProgram.py /data/$(TXT) --output /data/$$OUTPUT"; \
	if [ "$(USE_NLP)" = "1" ]; then \
		CMD="$$CMD --use-nlp"; \
	fi; \
	if [ "$(AGGRESSIVE)" = "1" ]; then \
		CMD="$$CMD --aggressive"; \
	fi; \
	if [ "$(BULLETS)" = "1" ]; then \
		CMD="$$CMD --preserve-bullets"; \
	fi; \
	docker-compose run --rm pdf-text-tools python /app/$$CMD; \
	echo "Preprocessed text saved to $$OUTPUT"

# Clean temporary files
clean:
	find . -name "*_cleaned.txt" -delete
	find . -name "*.pyc" -delete