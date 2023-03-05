{
 "cells": [

   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "# Python NLTK Natural Language Tool Kit\n",
    "## Topics covered in this video\n",
    "- Exploring the NLTK corpus  \n",
    "- Dictionary definitions   \n",
    "- Punctuation and stop words  \n",
    "- Stemming and lemmatization n",
    "- Sentence and word tokenizers \n",
    "- Parts of speech tagging  \n",
    "- word2vec  \n",
    "- Clustering and classifying\n",
    "\n",
    "### NLTK Setup\n",
    "First you need to install the nltk library with 'pip install nltk' or some equivalent shell command.  \n",
    "Then you need to download the nltk corpus by running  \n",
    "```python  \n",
    "import nltk  \n",
    "nltk.download()```\n",
    "This will open the NLTK downloader dialog window where you should just click Download All. The corpus is a large and varied body of sample documents that you'll need for this video, including dictionaries and word lists like stop words. You can uninstall it later if you have a shortage of disk space with *pip uninstall nltk*.\n"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 66,
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "<class 'nltk.text.Text'>\n",
      "260819\n",
      "19317\n",
      "['[', 'Moby', 'Dick', 'by', 'Herman', 'Melville', '1851', ']', 'ETYMOLOGY', '.']\n",
      "['[', 'Sense', 'and', 'Sensibility', 'by', 'Jane', 'Austen', '1811', ']', 'CHAPTER']\n"
     ]
    }
   ],
   "source": [
    "import nltk\n",
    "from nltk.book import *\n",
    "\n",
    "print(type(text1))\n",
    "print(len(text1))\n",
    "print(len(set(text1)))\n",
    "print(text1[:10])\n",
    "print(text2[:10])"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 67,
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "['austen-emma.txt', 'austen-persuasion.txt', 'austen-sense.txt', 'bible-kjv.txt', 'blake-poems.txt', 'bryant-stories.txt', 'burgess-busterbrown.txt', 'carroll-alice.txt', 'chesterton-ball.txt', 'chesterton-brown.txt', 'chesterton-thursday.txt', 'edgeworth-parents.txt', 'melville-moby_dick.txt', 'milton-paradise.txt', 'shakespeare-caesar.txt', 'shakespeare-hamlet.txt', 'shakespeare-macbeth.txt', 'whitman-leaves.txt']\n",
      "37360\n",
      "3106\n",
      "['What', 'say', 'you', '?']\n",
      "950\n"
     ]
    }
   ],
   "source": [
    "from nltk.corpus import gutenberg\n",
    "print(gutenberg.fileids())\n",
    "hamlet = gutenberg.words('shakespeare-hamlet.txt')\n",
    "print(len(hamlet))\n",
    "hamlet_sentences = gutenberg.sents('shakespeare-hamlet.txt')\n",
    "print(len(hamlet_sentences))\n",
    "print(hamlet_sentences[1024])\n",
    "print(len(gutenberg.paras('shakespeare-hamlet.txt')))"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### Get the count of a word in a document, or the context of every occurence of a word in a document."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 68,
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "26\n",
      "Displaying 7 of 7 matches:\n",
      "r him ,\" said I , now flying into a passion again at this unaccountable farrago\n",
      " employed in the celebration of the Passion of our Lord ; though in the Vision \n",
      "ce all mortal interests to that one passion ; nevertheless it may have been tha\n",
      "ing with the wildness of his ruling passion , yet were by no means incapable of\n",
      "it , however promissory of life and passion in the end , it is above all things\n",
      "o ' s lordly chest . So have I seen Passion and Vanity stamping the living magn\n",
      " Guernseyman , flying into a sudden passion . \" Oh ! keep cool -- cool ? yes , \n",
      "None\n",
      "Displaying 5 of 5 matches:\n",
      "one ,\" said Elinor , \" who has your passion for dead leaves .\" \" No ; my feelin\n",
      "r daughters , without extending the passion to her ; and Elinor had the satisfa\n",
      "r , if he was to be in the greatest passion !-- and Mr . Donavan thinks just th\n",
      "edness I could have borne , but her passion -- her malice -- At all events it m\n",
      "ling a sacrifice to an irresistible passion , as once she had fondly flattered \n",
      "None\n"
     ]
    }
   ],
   "source": [
    "print(text1.count('horse'))\n",
    "print(text1.concordance('passion'))\n",
    "print(text2.concordance('passion'))"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "**FreqDist and most_common**  \n",
    "We can use FreqDist to find the number of occurrences of each word in the text.  \n",
    "By getting len(vocab) we get the number of unique words in the text (including punctuation).  \n",
    "And we can get the most common words easily too."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 69,
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "19317\n",
      "[(',', 18713), ('the', 13721), ('.', 6862), ('of', 6536), ('and', 6024), ('a', 4569), ('to', 4542), (';', 4072), ('in', 3916), ('that', 2982), (\"'\", 2684), ('-', 2552), ('his', 2459), ('it', 2209), ('I', 2124), ('s', 1739), ('is', 1695), ('he', 1661), ('with', 1659), ('was', 1632)]\n"
     ]
    }
   ],
   "source": [
    "vocab = nltk.FreqDist(text1)\n",
    "print(len(vocab))\n",
    "print(vocab.most_common(20))"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "Here we got the 80 most common words, filtered only the ones with at least 3 characters, then sorted them descending by number of occurences.  \n",
    "A better way is to first remove all the *stop words* (see below), then get the FreqDist."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 70,
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "[('that', 2982), ('with', 1659), ('this', 1280), ('from', 1052), ('whale', 906), ('have', 760), ('there', 715), ('were', 680), ('which', 640), ('like', 624), ('their', 612), ('they', 586), ('some', 578), ('then', 571), ('when', 553), ('upon', 538), ('into', 520), ('ship', 507), ('more', 501), ('Ahab', 501), ('them', 471), ('what', 442), ('would', 421), ('been', 415), ('other', 412), ('over', 403)]\n"
     ]
    }
   ],
   "source": [
    "mc = sorted([w for w in vocab.most_common(80) if len(w[0]) > 3], key=lambda x: x[1], reverse=True)\n",
    "print(mc)"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### A dispersion plot shows you where in the document a word is used. You can pass in a list of words."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 71,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "image/png": "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\n",
      "text/plain": [
       "<Figure size 432x288 with 1 Axes>"
      ]
     },
     "metadata": {
      "needs_background": "light"
     },
     "output_type": "display_data"
    }
   ],
   "source": [
    "text1.dispersion_plot(['capture', 'whale', 'life', 'death', 'kill'])"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### Dictionary definitions\n",
    "Use wordnet synsets to get word definitions and examples of usage.  \n",
    "The [0] is required because synsets returns a list, with an entry for each POS."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 72,
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "unmitigated.a.01 - not diminished or moderated in intensity or severity; sometimes used as an intensifier\n",
      "['unmitigated suffering', 'an unmitigated horror', 'an unmitigated lie']\n"
     ]
    }
   ],
   "source": [
    "from nltk.corpus import wordnet as wn\n",
    "w = wn.synsets(\"unmitigated\")[0]\n",
    "print(w.name(), '-', w.definition())\n",
    "print(w.examples())"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### Punctuation and Stop Words\n",
    "Text analysis is often faster and easier if you can remove useless words.  \n",
    "NLTK provides a list of these stop words so it's easy to filter them out of your text prior to processing.  \n",
    "Here, 15% of our text is punctuation, and 40% is stop words. So we shrink the text by more than half by stripping out punctuation and stop words."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 73,
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~\n",
      "['i', 'me', 'my', 'myself', 'we', 'our', 'ours', 'ourselves', 'you', \"you're\", \"you've\", \"you'll\", \"you'd\", 'your', 'yours', 'yourself', 'yourselves', 'he', 'him', 'his', 'himself', 'she', \"she's\", 'her', 'hers', 'herself', 'it', \"it's\", 'its', 'itself', 'they', 'them', 'their', 'theirs', 'themselves', 'what', 'which', 'who', 'whom', 'this', 'that', \"that'll\", 'these', 'those', 'am', 'is', 'are', 'was', 'were', 'be', 'been', 'being', 'have', 'has', 'had', 'having', 'do', 'does', 'did', 'doing', 'a', 'an', 'the', 'and', 'but', 'if', 'or', 'because', 'as', 'until', 'while', 'of', 'at', 'by', 'for', 'with', 'about', 'against', 'between', 'into', 'through', 'during', 'before', 'after', 'above', 'below', 'to', 'from', 'up', 'down', 'in', 'out', 'on', 'off', 'over', 'under', 'again', 'further', 'then', 'once', 'here', 'there', 'when', 'where', 'why', 'how', 'all', 'any', 'both', 'each', 'few', 'more', 'most', 'other', 'some', 'such', 'no', 'nor', 'not', 'only', 'own', 'same', 'so', 'than', 'too', 'very', 's', 't', 'can', 'will', 'just', 'don', \"don't\", 'should', \"should've\", 'now', 'd', 'll', 'm', 'o', 're', 've', 'y', 'ain', 'aren', \"aren't\", 'couldn', \"couldn't\", 'didn', \"didn't\", 'doesn', \"doesn't\", 'hadn', \"hadn't\", 'hasn', \"hasn't\", 'haven', \"haven't\", 'isn', \"isn't\", 'ma', 'mightn', \"mightn't\", 'mustn', \"mustn't\", 'needn', \"needn't\", 'shan', \"shan't\", 'shouldn', \"shouldn't\", 'wasn', \"wasn't\", 'weren', \"weren't\", 'won', \"won't\", 'wouldn', \"wouldn't\"]\n",
      "260819\n",
      "221767\n",
      "122226\n"
     ]
    }
   ],
   "source": [
    "from string import punctuation\n",
    "print(punctuation)\n",
    "without_punct = [w for w in text1 if w not in punctuation]  # this is called a list comprehension\n",
    "\n",
    "from nltk.corpus import stopwords\n",
    "sw = stopwords.words('english')\n",
    "print(sw)\n",
    "without_sw = [w for w in without_punct if w not in sw] \n",
    "\n",
    "print(len(text1))\n",
    "print(len(without_punct))\n",
    "print(len(without_sw))"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### Stemming and Lemmatization\n",
    "These term normalization algorithms strip the word endings off to reduce the number of root words for easier matching.  \n",
    "This is useful for search term matching. [NLTK stemming docs](https://www.nltk.org/api/nltk.stem.html)"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 74,
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "is is\n",
      "are are\n",
      "bought bought\n",
      "buys buy\n",
      "giving give\n",
      "jumps jump\n",
      "jumped jump\n",
      "birds bird\n",
      "do do\n",
      "does doe\n",
      "did did\n",
      "doing do\n"
     ]
    }
   ],
   "source": [
    "from nltk.stem.porter import PorterStemmer\n",
    "st = PorterStemmer()\n",
    "words = ['is', 'are', 'bought', 'buys', 'giving', 'jumps', 'jumped', 'birds', 'do', 'does', 'did', 'doing']\n",
    "for word in words:\n",
    "    print(word, st.stem(word))"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "**WordNet Lemmatizer**   \n",
    "The difference is that the result of stemming may not be an actual word, but lemmatization returns the root word. NLTK supports both.  \n",
    "You can also try the Lancaster or Snowball stemmers. The Snowball stemmer supports numerous languages: Arabic, Danish, Dutch, English, Finnish, French, German, Hungarian, Italian, Norwegian, Portuguese, Romanian, Russian, Spanish and Swedish."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 75,
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "is is\n",
      "are are\n",
      "bought bought\n",
      "buys buy\n",
      "giving giving\n",
      "jumps jump\n",
      "jumped jumped\n",
      "birds bird\n",
      "do do\n",
      "does doe\n",
      "did did\n",
      "doing doing\n"
     ]
    }
   ],
   "source": [
    "from nltk.stem import WordNetLemmatizer\n",
    "wnl = WordNetLemmatizer()\n",
    "words = ['is', 'are', 'bought', 'buys', 'giving', 'jumps', 'jumped', 'birds', 'do', 'does', 'did', 'doing']\n",
    "for word in words:\n",
    "    print(word, wnl.lemmatize(word))"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### Sentence and Word Tokenizers\n",
    "Sentence tokenizer breaks text down into a list of sentences. It's pretty good at handling punctuation and decimal numbers.  \n",
    "[Word tokenizer](https://www.nltk.org/api/nltk.tokenize.html) breaks a string down into a list of words and punctuation.  \n",
    "It is also easy to get parts of speech using nltk.pos_tag. There are different tagsets, depending on how much detail you want. I like universal."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 76,
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "['Hello.', 'I am Joe!', 'I like Python.', '263.5 is a big number.']\n",
      "['The', 'quick', 'brown', 'fox', 'jumps', 'over', 'the', 'lazy', 'dog', '.']\n"
     ]
    }
   ],
   "source": [
    "from nltk.tokenize import sent_tokenize, word_tokenize\n",
    "s = 'Hello. I am Joe! I like Python. 263.5 is a big number.'  # 4 sentences\n",
    "print(sent_tokenize(s))\n",
    "\n",
    "w = word_tokenize('The quick brown fox jumps over the lazy dog.')\n",
    "print(w)"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### Parts of Speech Tagging\n",
    "To break a block of text down into its parts of speech use pos_tag.  \n",
    "The default tagset uses 2 or 3 letter tokens that are hard for me to understand. [StackOverflow](https://stackoverflow.com/questions/15388831/what-are-all-possible-pos-tags-of-nltk) has a great decoder for the default POS tags.  \n",
    "The Universal tagset gives a more familiar looking tag (noun, verb, adj).  \n",
    "NLTK includes several other tagsets you can try."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 77,
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "['The', 'quick', 'brown', 'fox', 'jumps', 'over', 'the', 'lazy', 'dog', '.']\n",
      "[('The', 'DT'), ('quick', 'JJ'), ('brown', 'NN'), ('fox', 'NN'), ('jumps', 'VBZ'), ('over', 'IN'), ('the', 'DT'), ('lazy', 'JJ'), ('dog', 'NN'), ('.', '.')]\n",
      "[('The', 'DET'), ('quick', 'ADJ'), ('brown', 'NOUN'), ('fox', 'NOUN'), ('jumps', 'VERB'), ('over', 'ADP'), ('the', 'DET'), ('lazy', 'ADJ'), ('dog', 'NOUN'), ('.', '.')]\n"
     ]
    }
   ],
   "source": [
    "w = word_tokenize('The quick brown fox jumps over the lazy dog.')\n",
    "print(w)\n",
    "print(nltk.pos_tag(w))\n",
    "print(nltk.pos_tag(w, tagset='universal'))"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### Word2Vec\n",
    "[Word2Vec](https://radimrehurek.com/gensim/models/word2vec.html) uses neural networks to analyze words in a corpus by using the contexts of words. \n",
    "It takes as its input a large corpus of text, and maps unique words to a vector space, such that \n",
    "words that share common contexts in the corpus are located in close proximity to one another in the space.  \n",
    "Word2Vec does NOT look at word meanings, it only finds words that are used in combination with other words. So *frying* and *pan* may have a high similarity.  \n",
    "You can see here the context of one word (pain) for two different corpora.  \n",
    "This uses the popular gensim library, which is not part of NLTK."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 78,
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "[('person', 0.9992861747741699), ('favourable', 0.998925507068634), ('meaning', 0.9987690448760986), ('effect', 0.9987439513206482), ('comfortable', 0.998741626739502), ('delay', 0.9987210035324097)]\n",
      "[('even', 0.9980006217956543), ('moment', 0.9979783296585083), ('hence', 0.9979231357574463), ('without', 0.9979217052459717), ('separate', 0.9979064464569092), ('Now', 0.9979038238525391)]\n"
     ]
    }
   ],
   "source": [
    "from gensim.models import Word2Vec\n",
    "emma_vec = Word2Vec(gutenberg.sents('austen-emma.txt'))\n",
    "leaves_vec = Word2Vec(gutenberg.sents('whitman-leaves.txt'))\n",
    "print(emma_vec.wv.most_similar('pain', topn=6))\n",
    "print(leaves_vec.wv.most_similar('pain', topn=6))"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 79,
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "30103\n",
      "[('mercy', 0.9284727573394775),\n",
      " ('liveth', 0.9062315821647644),\n",
      " ('truth', 0.8941866159439087),\n",
      " ('grace', 0.8938426375389099),\n",
      " ('glory', 0.8936725854873657),\n",
      " ('salvation', 0.8859732747077942),\n",
      " ('hosts', 0.8839103579521179),\n",
      " ('confession', 0.8837960958480835)]\n",
      "[('making', 0.9873884916305542),\n",
      " ('abundant', 0.9802387952804565),\n",
      " ('realm', 0.98007732629776),\n",
      " ('powers', 0.9798883199691772),\n",
      " ('twice', 0.9775580763816833)]\n"
     ]
    }
   ],
   "source": [
    "from gensim.models import Word2Vec\n",
    "from nltk.corpus import stopwords\n",
    "from string import punctuation\n",
    "import pprint as pp\n",
    "\n",
    "bible_sents = gutenberg.sents('bible-kjv.txt')\n",
    "sw = stopwords.words('english')\n",
    "bible = [[w.lower() for w in s if w not in punctuation and w not in sw] for s in bible_sents]\n",
    "print(len(bible))\n",
    "\n",
    "bible_vec = Word2Vec(bible)\n",
    "pp.pprint(bible_vec.wv.most_similar('god', topn=8))\n",
    "pp.pprint(bible_vec.wv.most_similar('creation', topn=5))"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### k-Means Clustering\n",
    "[Clustering](http://www.nltk.org/api/nltk.cluster.html) groups similar items together.  \n",
    "The K-means clusterer starts with k arbitrarily chosen means (or centroids) then assigns each vector to the cluster with the closest mean. It then recalculates the means of each cluster as the centroid of its vector members. This process repeats until the cluster memberships stabilize. [NLTK docs on this example](https://www.nltk.org/_modules/nltk/cluster/kmeans.html)  \n",
    "This example clusters int vectors, which you can think of as points on a plane. But you could also use clustering to cluster similar documents by vocabulary/topic."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 80,
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "k-means trial 0\n",
      "iteration\n",
      "iteration\n",
      "Clustered: [array([2, 1]), array([1, 3]), array([4, 7]), array([6, 7])]\n",
      "As: [0, 0, 1, 1]\n",
      "Means: [array([1.5, 2. ]), array([5., 7.])]\n"
     ]
    }
   ],
   "source": [
    "import numpy as np\n",
    "from nltk.cluster import KMeansClusterer, euclidean_distance\n",
    "\n",
    "vectors = [np.array(f) for f in [[2, 1], [1, 3], [4, 7], [6, 7]]]\n",
    "means = [[4, 3], [5, 5]]\n",
    "\n",
    "clusterer = KMeansClusterer(2, euclidean_distance, initial_means=means)\n",
    "clusters = clusterer.cluster(vectors, True, trace=True)\n",
    "\n",
    "print('Clustered:', vectors)\n",
    "print('As:', clusters)\n",
    "print('Means:', clusterer.means())"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "**k-Means Clustering, Example-2**  \n",
    "In this example we cluster an array of 6 points into 2 clusters.  \n",
    "The initial centroids are randomly chosen by the clusterer, and it does 10 iterations to regroup the clusters and recalculate centroids. "
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 103,
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "Clustered: [array([3, 3]), array([1, 2]), array([4, 2]), array([4, 0]), array([2, 3]), array([3, 1])]\n",
      "As: [0, 0, 1, 1, 0, 1]\n",
      "Means: [array([2.        , 2.66666667]), array([3.66666667, 1.        ])]\n",
      "classify([2 2]): 0\n"
     ]
    }
   ],
   "source": [
    "vectors = [np.array(f) for f in [[3, 3], [1, 2], [4, 2], [4, 0], [2, 3], [3, 1]]]\n",
    "\n",
    "# test k-means using 2 means, euclidean distance, and 10 trial clustering repetitions with random seeds\n",
    "clusterer = KMeansClusterer(2, euclidean_distance, repeats=10)\n",
    "clusters = clusterer.cluster(vectors, True)\n",
    "centroids = clusterer.means()\n",
    "print('Clustered:', vectors)\n",
    "print('As:', clusters)\n",
    "print('Means:', centroids)\n",
    "\n",
    "# classify a new vector\n",
    "vector = np.array([2,2])\n",
    "print('classify(%s):' % vector, end=' ')\n",
    "print(clusterer.classify(vector))"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "**Plot a Chart of the Clusters in Example-2**  \n",
    "Make a Scatter Plot of the two clusters using matplotlib.pyplot.   \n",
    "We plot all the points in cluster-0 blue, and all the points in cluster-1 red. Then we plot the two centroids in orange.  \n",
    "I used list comprehensions to create new lists for all the x0, y0, x1 and y1 values."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 104,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "image/png": "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\n",
      "text/plain": [
       "<Figure size 432x288 with 1 Axes>"
      ]
     },
     "metadata": {
      "needs_background": "light"
     },
     "output_type": "display_data"
    }
   ],
   "source": [
    "import matplotlib.pyplot as plt\n",
    "\n",
    "x0 = np.array([x[0] for idx, x in enumerate(vectors) if clusters[idx]==0])\n",
    "y0 = np.array([x[1] for idx, x in enumerate(vectors) if clusters[idx]==0])\n",
    "plt.scatter(x0,y0, color='blue')\n",
    "x1 = np.array([x[0] for idx, x in enumerate(vectors) if clusters[idx]==1])\n",
    "y1 = np.array([x[1] for idx, x in enumerate(vectors) if clusters[idx]==1])\n",
    "plt.scatter(x1,y1, color='red')\n",
    "\n",
    "xc = np.array([x[0] for x in centroids])\n",
    "yc = np.array([x[1] for x in centroids])\n",
    "plt.scatter(xc,yc, color='orange')\n",
    "plt.show()"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null
   "metadata":{}
   "outputs":[]
   "source": []
  }
 ],
 "metadata": {
  "kernelspec": {
   "display_name": "Python 3",
   "language": "python",
   "name": "python3"
  },
  "language_info": {
   "codemirror_mode": {
    "name": "ipython",
    "version": 3
   },
   "file_extension": ".py",
   "mimetype": "text/x-python",
   "name": "python",
   "nbconvert_exporter": "python",
   "pygments_lexer": "ipython3",
   "version": "3.7.0"
  }
 },
 "nbformat": 4,
 "nbformat_minor": 2
}