# AI-Powered Presentation Generator

This project uses **LangChain**, **Groq LLM**, and **Tavily Search** to automatically generate professional PowerPoint presentations from any topic.

## Features
- Takes a topic as input  
- Searches the web for up-to-date information (via Tavily)  
- Summarizes the topic into a clear, structured outline using LLM  
- Automatically generates a **PowerPoint (.pptx)** with styled slides  

## Requirements
- Python **3.9+**  
- Install dependencies:  
  ```bash
  pip install -r requirements.txt
## Environment Variables

Create a `.env` file in the project root with your API keys:  

```env
  GROQ_API_KEY=your_groq_key_here
  TAVILY_API_KEY=your_tavily_key_here 


