<template>
  <v-app id="app">
    <v-app-bar
      app
      color="primary"
      dark
      class="d-flex justify-center"
    >
      <div class="d-flex align-center">
        <h3>AMAZON KEYWORDS ANALYZER</h3>
      </div>

    </v-app-bar>

    <v-main>
      <MainComponent/>
    </v-main>
  </v-app>
</template>

<script>
import MainComponent from './components/MainComponent';

export default {
  name: 'App',

  components: {
    MainComponent,
  },

};
</script>
<style>
#app{
 font-family: 'Roboto', sans-serif;
}
@import url('https://fonts.googleapis.com/css2?family=Roboto&display=swap');
</style>