package api.endpoinds;

/*
Here we will maintain only URL'S of PetStore API :-
=====================================================================
Swagger URI          : https://petstore.swagger.io/
Create user (Post)   : https://petstore.swagger.io/v2/user
Get user (Get)       : https://petstore.swagger.io/v2/user/{username}
Update user (put)    : https://petstore.swagger.io/v2/user/{username}
Delete user (Delete) : https://petstore.swagger.io/v2/user/{username}
*/

public class Routes {

	public static String base_url = "https://petstore.swagger.io/v2";
	// USER MODEL'S URL FOR ALL HTTP REQUEST -
	public static String post_url = base_url + "/user";
	public static String get_url = base_url + "/user/{username}";
	public static String put_url = base_url + "/user/{username}";
	public static String delete_url = base_url + "/user/{username}";
	
	// STORE MODEL'S URL FOR ALL HTTP REQUEST -
	
	// PET MODEL'S URL FOR ALL HTTP REQUEST -
}
