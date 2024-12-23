import streamlit as st
import pickle
import pandas as pd

def recommend(user_id, genres):
    # Calculate cosine similarity between users
    # Find most similar user
    x = list(sorted(enumerate(c_similarity[user_id - 1]), key=lambda x: x[1], reverse=True)[1:2])
    
    # Get items rated by both users
    user1_items = set(df[df['user id'] == user_id]['item id']) 
    user2_items = set(df[df['user id'] == x[0][0] + 1]['item id'])

    print('Recommend Movie:')
    
    # Find movies recommended to similar user but not rated by the current user
    intersection_movies = user2_items - user1_items
    recommend_movies_id = []
    for j in intersection_movies:
        for i in new_item['genres'].values[j - 1]:
            if i == genres:
                recommend_movies_id.append(j)

    print(recommend_movies_id)
    movie = recommend_movies_id
    count = 0
    
    # Extract ratings for recommended movies
    rating = []
    for i in movie:
        rating.append(df[df['item id'] == i]["rating"].values[0])

    # Combine ratings with movie IDs
    new_list = [(x, y) for x, y in zip(rating, movie)]
    movies_rating = sorted(new_list, reverse=True)
    movie_titles = []  # Initialize an empty list to accumulate movie titles
    for rating, movie_id in movies_rating:
        movie_title = item[item['movie id'] == movie_id]['movie title'].values[0]
        movie_titles.append(movie_title)  # Append movie title to the list
        if count < 4:  # Print only top 5 recommendations
            count += 1
        else:
            break  # Exit the loop if count exceeds 4
    return movie_titles  # Return the list of movie titles

# Load data
movie_dict = pickle.load(open('item_dict.pkl', 'rb'))
item = pd.DataFrame(movie_dict)

n_movie_dict = pickle.load(open('new_item_dict.pkl', 'rb'))
new_item = pd.DataFrame(n_movie_dict)

df_dict = pickle.load(open('df_dict.pkl', 'rb'))
df = pd.DataFrame(df_dict)  

c_similarity = pickle.load(open('c_similarity.pkl', 'rb'))

st.title("Recommend System")
user_id = list(range(1, 944))
selected_user_id = st.selectbox(
    'Select User id:',
    user_id
)

gen = [ 'Action', 'Adventure', 'Animation', "Children's", 'Comedy', 'Crime', 'Documentary', 'Drama', 'Fantasy', 'Film-Noir', 'Horror', 'Musical', 'Mystery', 'Romance', 'Sci-Fi', 'Thriller', 'War', 'Western','unknown']
selected_genres = st.selectbox(
    'Select genres:',
    gen
)

if st.button('Recommend'):
    recommendations = recommend(selected_user_id, selected_genres)
    for recommendation in recommendations:
        st.write(recommendation)
