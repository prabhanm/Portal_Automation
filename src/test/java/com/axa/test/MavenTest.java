package com.axa.test;


import static org.junit.Assert.assertTrue;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;

import org.axa.framework.CommonFunctions;
import org.axa.portal.page.utility;
import org.junit.Test;

import io.qameta.allure.Description;
import io.qameta.allure.Story;

public class MavenTest {
	
	@Test
	@Story("Test story")
	@Description("This is the allure testing case1")
	 public void test1() throws ClassNotFoundException, NoSuchMethodException, SecurityException, IllegalAccessException, IllegalArgumentException, InvocationTargetException, InstantiationException, IOException
	    {
         System.out.println(CommonFunctions.getCurrentDate("HH:mm"));
		
		
		//BC_utility util=new BC_utility();
		//util.BC_methodToInokeFunction();
		
		utility util=new utility(); 
		util.methodToInokeFunction();
	    }
	
	
	@Test
	@Story("Test story")
	@Description("This is the allure testing case1")
    public void shouldAnswerWithTrue()
    {
		System.out.println("This Test2 package");
        assertTrue( true );
    }

}
