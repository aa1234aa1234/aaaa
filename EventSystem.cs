using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;

namespace ohmygod
{
    class Event
    {
        public string EventId;

        public Event(string id) { EventId = id; }
    }

    class Listener
    {
        public EventFunction eventFunction;

        public delegate void EventFunction(params object[] args);

        public Listener(EventFunction eventFunction) { this.eventFunction = eventFunction; }
    }

    class EventListener
    {
        protected string EventId;
        public EventListener(string eventId)
        {
            EventId = eventId;
            EventSystem.GetInstance().Subscribe(EventId, new Listener((params object[] args) => { OnEvent(args); }));
        }

        public virtual void OnEvent(params object[] args) { }
    }

    internal class EventSystem
    {
        private static EventSystem instance;

        private Dictionary<string, List<Listener>> listeners=new();

        public EventSystem() { }
        ~EventSystem() { instance = null; }

        public static EventSystem GetInstance() { if (instance == null) instance = new EventSystem(); return instance; }

        public void Subscribe(string eventName, Listener listener) { 
            if(!listeners.ContainsKey(eventName)) listeners.Add(eventName, new List<Listener>());
            listeners[eventName].Add(listener); 
        }

        public void Unsubscribe(string eventName, Listener listener)
        {
            listeners[eventName].Remove(listener);
        }

        public void DispatchEvent(Event evnt, params object[] args)
        {
            List<Listener> list = listeners[evnt.EventId];
            foreach (Listener listener in list)
            {
                listener.eventFunction(args);
            }
        }
    }
}
