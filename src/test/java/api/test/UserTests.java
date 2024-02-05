package api.test;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.testng.Assert;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import com.github.javafaker.Faker;

import api.endpoinds.UserEndPoints;
import api.payload.User;
import io.restassured.response.Response;
import jdk.internal.org.jline.utils.Log;

public class UserTests {
	
	Faker faker;
	User userPayload;
	
	@BeforeClass
	public void setupData() {
		faker=new Faker();
		userPayload=new User();
		//Logger Log;
		
		userPayload.setId(faker.idNumber().hashCode());
		userPayload.setUsername(faker.name().username());
		userPayload.setFirstName(faker.name().firstName());
		userPayload.setLastName(faker.name().lastName());
		userPayload.setEmail(faker.internet().safeEmailAddress());
		userPayload.setPassword(faker.internet().password());
		userPayload.setPhone(faker.phoneNumber().cellPhone());
		
		// To generate the logs
		//Log=LogManager.getLogger(this.getClass());
	}
	@Test(priority = 1)
	public void testPostUser(){
		//Log.info("************ Creating user**************");
		Response response=UserEndPoints.createUser(userPayload);
		response.then().log().all();
		Assert.assertEquals(response.getStatusCode(),200);
		//Log.info("*********** User is created *********");
	}
	@Test(priority = 2)
	public void testGetUserUserByName() {
		//Log.info("************ Reading User Info **************");
		Response response=UserEndPoints.readUser(this.userPayload.getUsername());
		response.then().log().all();
		Assert.assertEquals(response.getStatusCode(),200);
		//Log.info("************ User info is displayed **************");
	}
	@Test(priority = 3)
	public void testUpdateUserByName() {
		
		// Update data using payload
		//Log.info("************ Updating User Info**************");
		userPayload.setFirstName(faker.name().firstName());
		userPayload.setLastName(faker.name().lastName());
		userPayload.setEmail(faker.internet().safeEmailAddress());
		
		Response response=UserEndPoints.updateUser(this.userPayload.getUsername(), userPayload);
		response.then().log().all();
		Assert.assertEquals(response.getStatusCode(),200);
		//Log.info("************ User Is Updated **************");
		
		// Checking data after updated
		
		Response responseAfterUpdate=UserEndPoints.updateUser(this.userPayload.getUsername(), userPayload);
		response.then().log().all();
		Assert.assertEquals(responseAfterUpdate.getStatusCode(),200);
	}
	@Test(priority = 4)
	public void deleteUserByName() {
		//Log.info("************ Deleting User **************");
		Response response=UserEndPoints.deleteUser(this.userPayload.getUsername());
		response.then().log().all();
		Assert.assertEquals(response.getStatusCode(), 200);
		//Log.info("************ User Is Deleted **************");
	}

}
