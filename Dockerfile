FROM python:3.8.3-slim

RUN apt-get upgrade && apt-get update \
    && pip3 install --upgrade pip \
    && pip install pandas xlrd==1.2.0 openpyxl \
    && pip install numpy

RUN groupadd -g 1000 app \
    && useradd -u 1000 -g app -s /bin/bash -m app \
    && usermod -aG sudo app \
    && echo 'app    ALL=(ALL) NOPASSWD:ALL' >> /etc/sudoers \
    # && mkdir -p /usr/src/app \
    && chown -R app:app /usr/src 
    # && chmod -R 777 /usr/src/app

WORKDIR /app
COPY . /app/

CMD /bin/bash

# docker run -it -v <source>:<target> <containerID>
# python main.py <inputFile>