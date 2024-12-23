import streamlit as st
import pickle
import pandas as pd

def recommend(movie):
    movie_index=movies[movies['movie title']==movie].index[0]
    distance=similarity[movie_index]
    movie_list=list(enumerate(distance))
    for j in range(0,len(mean_ratings)):
        movie_list[j]=movie_list[j]+(mean_ratings[j],)
    sorted_list = sorted(movie_list, key=lambda x: (-x[1],-x[2]))[1:11]
    
    recommend_movie=[]
    for i in sorted_list:
        recommend_movie.append(movies.iloc[i[0],1])
    return recommend_movie

movie_dict=pickle.load(open('movies_dict.pkl','rb'))
movies=pd.DataFrame(movie_dict)

similarity=pickle.load(open('similarity.pkl','rb'))
mean_ratings=pickle.load(open('mean_rating.pkl','rb'))
st.title("Recommend System")

selected_movie_name=st.selectbox(
'Select previus movie you watched:',
movies['movie title'].values )

if st.button('Recommend'):
    recoomend=recommend(selected_movie_name)
    for i in recoomend:
        st.write(i)