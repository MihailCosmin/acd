import os
import sys
import base64
import regex

from os.path import basename
import json
from datetime import datetime
from ollama import Client

from time import sleep
from tqdm import tqdm

from traceback import format_exc

if __name__ == "__main__":
    from filelist import list_files
    from filepath import clean_path
else:
    from .filelist import list_files
    from .filepath import clean_path

def ollama_with_files(
    prompt: str,
    base_url: str = "http://10.3.1.174:11434",
    model: str = "gpt-oss:120b",
    files: list[str] = None,
    num_predict: int = 12,          # cap tokens for short answers
    temperature: float = 0.1,       # keep it concise/consistent
    keep_alive: str = "5m",         # keep model hot during batches
) -> str:
    # Build context from files
    file_context = "\n<files_context>"
    for file in files or []:
        with open(clean_path(file), "r", encoding="utf-8") as f:
            file_context += (
                f"\n<file name='{os.path.basename(file)}'>\n"
                + f.read()
                + "\n</file>\n"
            )
    file_context += "</files_context>"

    # Call Ollama
    response = Client(host=base_url).generate(
        model=model,
        prompt=prompt + file_context,
        keep_alive=keep_alive,
        options={
            "num_predict": num_predict,
            "temperature": temperature,
        },
    )

    return (
        response["response"]
        .strip()
        .replace("```xml", "")
        .replace("```", "")
        .strip()
    )

def ollama_with_images(prompt: str, host: str = "http://10.3.1.174:11434/api", model: str = "gemma3:27b-it-q8_0", images: list = None):
    client = Client(host=host)

    resp = client.chat(
        model=model,
        messages=[
            {
                'role': 'user',
                'content': prompt,
                'images': images
            }
        ]
    )

    return resp['message']['content']


if __name__ == "__main__":
    # results = {}
    # prompts = 0
    # for modification_document in tqdm(list_files(
    #         r"C:\Users\munte\Downloads\Dubai Air Wing\Work\777 AIPC\short test",
    #         True,
    #         [".txt", ".TXT"]), desc="Processing modification documents", colour="green"):
    #     results[basename(modification_document)] = {}
    #     for task in tqdm(list_files(
    #             r"C:\Users\munte\Downloads\Dubai Air Wing\Work\777 AIPC\777 - Supplement GTI-1435-IPCS Rev IR",
    #             True,
    #             [".txt", ".TXT"]), desc="Processing tasks", colour="blue"):
    
    #         # modification_document = r"C:\Users\munte\Downloads\Dubai Air Wing\Work\777 AIPC\short test\10079-01-A-INST-F01-R00.txt"
    #         # task = r"C:\Users\munte\Downloads\Dubai Air Wing\Work\777 AIPC\777 - Supplement GTI-1435-IPCS Rev IR\Chapter 25\25-21\25-21-01\25-21-01-01\25-21-01-01 - Forward Entry Ceiling Panels - Install.txt"
    #         result = ollama_with_files(
    #                     prompt=f"Is the modification document: {basename(modification_document)} applicable" +
    #                            f" to the task: {basename(task)}?" +
    #                            "IMPORTANT: Only answer with Yes or No and a percentage of how sure you are of that answer!" +
    #                            "Ex 1: Yes, 80%" +
    #                            "Ex 2: No, 50%",
    #                     model="gpt-oss:latest",  # Cosmin: gpt-oss:latest is 6-7 times faster than gpt-oss:120b
    #                     files=[
    #                         modification_document,
    #                         task
    #                     ]
    #                 )
    #         prompts += 1
    #         results[basename(modification_document)][basename(task)] = result
    # print(f"Total prompts sent: {prompts}")

    # with open(clean_path(r"C:\Users\munte\Downloads\Dubai Air Wing\Work\777 AIPC\AI Check results.json"), "w", encoding="utf-8") as f:
    #     json.dump(results, f, ensure_ascii=False, indent=4)

    modification_document = r"C:\Users\munte\Downloads\Dubai Air Wing\Work\777 AIPC\short test\10079-01-A-INST-F01-R00.txt"
    task = r"C:\Users\munte\Downloads\Dubai Air Wing\Work\777 AIPC\777 - Supplement GTI-1435-IPCS Rev IR\Chapter 25\25-21\25-21-01\25-21-01-01\25-21-01-01 - Forward Entry Ceiling Panels - Install.txt"
    result = ollama_with_files(
                prompt=f"Is the modification document: {basename(modification_document)} applicable" +
                        f" to the task: {basename(task)}?" +
                        "IMPORTANT: Only answer with Yes or No and a percentage of how sure you are of that answer!" +
                        "Ex 1: Yes, 80%" +
                        "Ex 2: No, 50%",
                model="gpt-oss:latest",  # Cosmin: gpt-oss:latest is 6-7 times faster than gpt-oss:120b
                files=[
                    modification_document,
                    task
                ]
            )
    print(result)
    