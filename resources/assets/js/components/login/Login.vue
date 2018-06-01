<template>
    <v-container>
  <v-form @submit.prevent="login">
    <v-text-field
      v-model="form.email"
      label="E-mail"
      type="email"
      required
    ></v-text-field>

     <v-text-field
      v-model="form.password"
      type="password"
      label="Password"
      required
    ></v-text-field>

    <v-btn
     color="green"
     type="submit"
     >Login</v-btn>
  </v-form>
  </v-container>
</template>
<script>
export default {
    data(){
        return {
            form :{
                email:null,
                passowrd:null
            }
        }
    },
    methods:{
        login(){
            User.login(this.form)  

            
        }
    }

}
</script>

<style>

</style>
