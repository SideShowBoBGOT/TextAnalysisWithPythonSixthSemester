import nltk
with open('traindata.txt', 'r') as file:
    text = ''.join(file.readlines())
sentences = nltk.tokenize.sent_tokenize(text)
import spacy
nlp = spacy.load('en_core_web_lg')
trainset = []
for sent in sentences:
    doc = nlp(sent)
    entities = []
    for ent in doc.ents:
        entities.append(
            (
                ent.start_char,
                ent.end_char,
                ent.label_
            )
        )
    trainset.append((sent, {'entities': entities}))
format(trainset)
with open('prepared_train_data.txt', 'w') as file:
    file.write(format(trainset))