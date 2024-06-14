import tensorflow as tf
from tensorflow import keras

def load_model(model_name):
    return tf.keras.models.load_model(
        'model/{}.keras'.format(model_name),
        )

# Models
model_name_tc = "tc_index_kt_b2.2_2024_02_02_20_08_42_190891"
model_name_wic = "wic_index_kt_b2.4_2024_02_02_20_08_42_190891"
model_name_ac = "ac_index_kt_b2.4_2024_02_18_05_42_21_089087"

best_model_tc = load_model(model_name_tc)
best_model_wic = load_model(model_name_wic)
best_model_ac = load_model(model_name_ac)

print(best_model_tc)
print(best_model_wic)
print(best_model_ac)