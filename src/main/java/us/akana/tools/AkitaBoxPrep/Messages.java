package us.akana.tools.AkitaBoxPrep;


import java.util.MissingResourceException;
import java.util.ResourceBundle;

/**
 * Used to get Strings from messages.properties
 * @author Jaden Unruh
 */
public class Messages {

	/**
	 * A bundle representing all of the Strings
	 */
	private static final ResourceBundle RESOURCE_BUNDLE = ResourceBundle.getBundle("us.akana.tools.AkitaBoxPrep.messages");

	/**
	 * Gets the specified string from {@link Messages#RESOURCE_BUNDLE}
	 * @param key the key of the String we need
	 * @return the string itself
	 */
	public static String getString(String key) {
		try {
			return RESOURCE_BUNDLE.getString(key);
		} catch (MissingResourceException e) {
			return '!' + key + '!';
		}
	}
}
