plugins {
    id("java")
    id("com.github.johnrengelman.shadow") version "8.1.1"
}

group = "org.example"
version = "1.0-SNAPSHOT"

repositories {
    mavenCentral()
}

dependencies {
    implementation("org.apache.poi:poi:5.3.0")
    implementation("org.apache.poi:poi-ooxml:5.3.0")
    implementation("org.apache.logging.log4j:log4j-api:2.24.1")
    implementation("org.apache.logging.log4j:log4j-core:2.24.1")
    //testImplementation(platform("org.junit:junit-bom:5.10.0"))
    //testImplementation("org.junit.jupiter:junit-jupiter")
}

tasks.withType<Jar> {
    manifest {
        attributes["Main-Class"] = "org.example.Main"
    }
    //from(configurations.compileClasspath.get().map { if (it.isDirectory()) it else zipTree(it) })
}


/*
tasks.test {
    useJUnitPlatform()
}*/
