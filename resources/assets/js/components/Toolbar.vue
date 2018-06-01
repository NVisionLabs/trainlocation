<template>
  <v-toolbar>
  <!--  <v-toolbar-side-icon></v-toolbar-side-icon> -->
    <v-toolbar-title>Train Location</v-toolbar-title>
    <v-spacer></v-spacer>
    <div class="hidden-sm-and-down">
      <router-link to="/">
        <v-btn flat>Location</v-btn>
      </router-link>
      <router-link to="/login"> 
      <v-btn flat>Login</v-btn>
      </router-link>
    </div>
  </v-toolbar>
</template>

<script>
export default {

}
</script>

<style>

</style>
