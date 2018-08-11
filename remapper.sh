#!/usr/bin/env bash

cp /Users/bbudhathoki/Desktop/GInputParam/*.csv /Users/bbudhathoki/_Code/_Client/gemini/utils/resource/input/
source ve/bin/activate
python ./remapper.py
cp /Users/bbudhathoki/_Code/_Client/gemini/utils/resource/output/*.json /Users/bbudhathoki/_Code/_Client/gemini/GEnergyOptimizer/GEnergyOptimizer/Model/DataObject/Form/Config/
