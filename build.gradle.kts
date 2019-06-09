import org.jetbrains.kotlin.gradle.tasks.KotlinCompile

plugins {
    kotlin("jvm") version "1.3.11"
}

group = "com.jamesstapleton.autofillwordform"
version = "1.0-SNAPSHOT"

repositories {
    mavenCentral()
}

dependencies {
    compile(kotlin("stdlib-jdk8"))
    compile(group = "org.apache.poi", name = "poi-ooxml", version = "4.1.0")
}

tasks.withType<KotlinCompile> {
    kotlinOptions.jvmTarget = "1.8"
}

tasks.withType<Jar> {
    manifest {
        attributes["Main-Class"] = "com.jamesstapleton.autofillwordform.app.MainKt"
    }

    from(configurations.runtime.map { if (it.isDirectory) it else zipTree(it) })

    doLast {
        val execJarFile = File(archivePath.parentFile, "${archivePath.nameWithoutExtension}-exec.jar")
        val scriptFile = File(projectDir, "src/main/scripts/launcher-stub.sh").readText()
        execJarFile.outputStream().use {
            it.write(scriptFile.toByteArray())
            it.write("\n".toByteArray())
            archivePath.inputStream().use { inStream -> inStream.copyTo(it) }
        }
        execJarFile.setExecutable(true)
        execJarFile.renameTo(archivePath)
    }
}
