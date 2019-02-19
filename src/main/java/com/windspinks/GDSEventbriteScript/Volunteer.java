package com.windspinks.GDSEventbriteScript;

public class Volunteer {
    private String firstName, lastName, homeAddress, homeAddressTwo, city, state, zipCode, cellPhone, visitor, firstTime, age, firstChoice, secondChoice, thirdChoice, emailAddress, special;

    public Volunteer() {
    }

    public Volunteer(String firstName, String lastName, String homeAddress, String homeAddressTwo, String city, String state, String zipCode, String cellPhone, String visitor, String firstTime, String age, String firstChoice, String secondChoice, String thirdChoice, String emailAddress, String special) {
        this.firstName = firstName;
        this.lastName = lastName;
        this.homeAddress = homeAddress;
        this.homeAddressTwo = homeAddressTwo;
        this.city = city;
        this.state = state;
        this.zipCode = zipCode;
        this.cellPhone = cellPhone;
        this.visitor = visitor;
        this.firstTime = firstTime;
        this.age = age;
        this.firstChoice = firstChoice;
        this.secondChoice = secondChoice;
        this.thirdChoice = thirdChoice;
        this.emailAddress = emailAddress;
        this.special = special;
    }

    public String getFirstName() {
        return firstName;
    }

    public void setFirstName(String firstName) {
        this.firstName = firstName;
    }

    public String getLastName() {
        return lastName;
    }

    public void setLastName(String lastName) {
        this.lastName = lastName;
    }

    public String getHomeAddress() {
        return homeAddress;
    }

    public void setHomeAddress(String homeAddress) {
        this.homeAddress = homeAddress;
    }

    public String getHomeAddressTwo() {
        return homeAddressTwo;
    }

    public void setHomeAddressTwo(String homeAddressTwo) {
        if (homeAddressTwo.equals("Apartment, suite, unit etc. (optional)") || homeAddressTwo.equals("None")){
            this.homeAddressTwo = "";
        } else {
            this.homeAddressTwo = homeAddressTwo;
        }
    }

    public String getCity() {
        return city;
    }

    public void setCity(String city) {
        this.city = city;
    }

    public String getState() {
        return state;
    }

    public void setState(String state) {
        this.state = state;
    }

    public String getZipCode() {
        return zipCode;
    }

    public void setZipCode(String zipCode) {
        this.zipCode = zipCode;
    }

    public String getCellPhone() {
        return cellPhone;
    }

    public void setCellPhone(String cellPhone) {
        if (!cellPhone.contains("-")){
            cellPhone = cellPhone.replaceFirst("(\\d{3})(\\d{3})(\\d+)", "$1-$2-$3");
        }
        if (cellPhone.contains("(")) {
            cellPhone = cellPhone.replace("(", "");
        }
        if (cellPhone.contains(")")) {
            cellPhone = cellPhone.replace(")", "");
        }
        this.cellPhone = cellPhone;
    }

    public String getVisitor() {
        return visitor;
    }

    public void setVisitor(String visitor) {
        this.visitor = visitor;
    }

    public String getFirstTime() {
        return firstTime;
    }

    public void setFirstTime(String firstTime) {
        this.firstTime = firstTime;
    }

    public String getAge() {
        return age;
    }

    public void setAge(String age) {
        this.age = age;
    }

    public String getFirstChoice() {
        return firstChoice;
    }

    public void setFirstChoice(String firstChoice) {
        this.firstChoice = firstChoice;
    }

    public String getSecondChoice() {
        return secondChoice;
    }

    public void setSecondChoice(String secondChoice) {
        this.secondChoice = secondChoice;
    }

    public String getThirdChoice() {
        return thirdChoice;
    }

    public void setThirdChoice(String thirdChoice) {
        this.thirdChoice = thirdChoice;
    }

    public String getEmailAddress() {
        return emailAddress;
    }

    public void setEmailAddress(String emailAddress) {
        this.emailAddress = emailAddress;
    }

    public String getSpecial() {
        return special;
    }

    public void setSpecial(String special) {
        if (special.toLowerCase().equals("no") || special.toLowerCase().equals("none")){} else {
        this.special = special;
        }
    }

    public void trimChoices() {
        if (firstChoice != null && !firstChoice.equals("")) {
            firstChoice = firstChoice.substring(0, firstChoice.indexOf("."));
        }
        if (secondChoice != null && !secondChoice.equals("")) {
            secondChoice = secondChoice.substring(0, secondChoice.indexOf("."));
        }
        if (thirdChoice != null && !thirdChoice.equals("")) {
            thirdChoice = thirdChoice.substring(0, thirdChoice.indexOf("."));
        }

    }

    @Override
    public String toString() {
        return "Volunteer{" +
            "firstName='" + firstName + '\'' +
            ", lastName='" + lastName + '\'' +
            ", homeAddress='" + homeAddress + '\'' +
            ", homeAddressTwo='" + homeAddressTwo + '\'' +
            ", city='" + city + '\'' +
            ", state='" + state + '\'' +
            ", zipCode='" + zipCode + '\'' +
            ", cellPhone='" + cellPhone + '\'' +
            ", visitor='" + visitor + '\'' +
            ", firstTime='" + firstTime + '\'' +
            ", age='" + age + '\'' +
            ", firstChoice='" + firstChoice + '\'' +
            ", secondChoice='" + secondChoice + '\'' +
            ", thirdChoice='" + thirdChoice + '\'' +
            ", emailAddress='" + emailAddress + '\'' +
            ", special='" + special + '\'' +
            '}';
    }
}


//    City	ST	ZIP	Cell Phone	V?	1?	Age	1ST	2ND	3RD
//email, special