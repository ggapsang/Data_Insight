import os
from datetime import datetime
import json

import numpy as np
import pandas as pd

import tensorflow as tf
from tensorflow.keras.preprocessing.text import Tokenizer
from tensorflow.keras.preprocessing.sequence import pad_sequences
# from tensorflow.keras.models import Sequential
# from tensorflow.keras.layers import Embedding, Bidirectional, LSTM, TimeDistributed, Dense, concatenate

from tensorflow.keras.models import Model, load_model
# from keras.layers import Embedding, Conv1D, GlobalMaxPooling1D, Dense, Reshape, Lambda
from keras.preprocessing.text import tokenizer_from_json
# from tensorflow.keras.utils import to_categorical
# from tensorflow.keras import layers


class Oracle() :
    def __init__(self) :
        pass
    def predict_text(self, texts, model, tokenizer, batch=True, print_pad=False) :
        if batch :
            sequences_texts = tokenizer.texts_to_sequences(texts)
            padded_sequences_texts = pad_sequences(sequences_texts, maxlen=36, padding='post')
            predictions = model.predict(padded_sequences_texts)
            predicted_classes = np.argmax(predictions, axis=-1)
            padded_predicted_classes = [''.join(map(str, lst)) for lst in predicted_classes.tolist()]

            return padded_predicted_classes

        else :
            sequences_text = tokenizer.texts_to_sequences([texts])
            padded_sequences_text = pad_sequences(sequences_text, maxlen=36, padding='post')

            if print_pad :
                print(padded_sequences_text)

            # 모델 예측
            predictions = model.predict(padded_sequences_text)
            predicted_classes = np.argmax(predictions, axis=-1)

            padded_predicted_classes = [''.join(map(str, lst)) for lst in predicted_classes.tolist()]

            return padded_predicted_classes

    def predict_in_df(self, model_name, tokenizer_name, inference_data, df_return=True) :
        """예측을 수행한다"""
        
        model = load_model(model_name, compile=False)
        
        with open(tokenizer_name) as f:
            data = json.load(f)
            tokenizer = tokenizer_from_json(data)

        lst_tag_no = inference_data['TAG NO'].tolist()
        predictions = self.predict_text(lst_tag_no, model, tokenizer, batch=True)
        inference_data['prediction'] = predictions
        inference_data['prediction'] = inference_data.apply(lambda row : (row['prediction'][:len(row['TAG NO'])]), axis=1)

        return inference_data
